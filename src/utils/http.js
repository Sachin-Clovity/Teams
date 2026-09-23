const MAX_RETRIES = 3;
const BASE_DELAY_MS = 500;
const MAX_DELAY_MS = 4000;

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Retries a fetch when the server signals it's temporarily overloaded (429 rate-limited, or a
// 5xx it might recover from) instead of dropping the request on the first hiccup — which is
// what every webhook/Graph call in this app used to do silently. Honors the server's own
// Retry-After header when it sends one (Power Automate and Graph both do this for 429s)
// instead of guessing a delay; otherwise backs off exponentially, capped so one invocation
// can't run long enough to hit Forge's own execution time limit.
//
// A burst of many events at once (e.g. a bulk issue import firing hundreds of trigger
// invocations back to back) will still overwhelm this — retrying every one of them just shifts
// when they fail, it doesn't create capacity that isn't there. This only absorbs the ordinary
// case of a handful of requests briefly tripping a limit, and — just as importantly — always
// logs plainly when it gives up, so a real failure is visible in `forge logs` instead of a
// silent gap where a notification should have been.
export async function fetchWithRetry(fetch, url, options = {}, label = 'request') {
  let lastRes;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    lastRes = await fetch(url, options);
    if (lastRes.ok) return lastRes;
    // Only 429 and 5xx are worth retrying — a 4xx like 400/401/404 will fail the same way
    // every time, so retrying it would just waste the invocation's time budget.
    if (lastRes.status !== 429 && lastRes.status < 500) return lastRes;
    if (attempt === MAX_RETRIES) break;

    let delayMs = BASE_DELAY_MS * 2 ** attempt;
    try {
      const retryAfter = lastRes.headers?.get?.('Retry-After');
      if (retryAfter) delayMs = parseInt(retryAfter, 10) * 1000;
    } catch { /* header not readable — fall back to the exponential guess above */ }
    delayMs = Math.min(delayMs, MAX_DELAY_MS);

    console.log(`[retry] ${label} got ${lastRes.status}, retrying in ${delayMs}ms (attempt ${attempt + 1}/${MAX_RETRIES})`);
    await sleep(delayMs);
  }
  console.log(`[retry] ${label} still failing after ${MAX_RETRIES} retries — giving up, last status: ${lastRes.status}`);
  return lastRes;
}
