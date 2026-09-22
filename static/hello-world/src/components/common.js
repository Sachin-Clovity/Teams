import React, { createContext, useCallback, useContext, useState } from 'react';

// ── Shared page chrome — used by all four Jira-embedded pages so they read as one
// consistent app instead of four independently-built forms. ─────────────────────

// ── Toasts — transient action feedback (save succeeded, test failed) that clears itself,
// as opposed to alert-err/alert-ok banners which are for a persistent page-level problem
// (e.g. "failed to load") that should stay visible until the user does something about it. ──
const ToastContext = createContext(() => {});

export function ToastProvider({ children }) {
  const [toasts, setToasts] = useState([]);

  const dismiss = useCallback(id => {
    setToasts(prev => prev.filter(t => t.id !== id));
  }, []);

  const showToast = useCallback((message, type = 'success') => {
    const id = `${Date.now()}-${Math.random()}`;
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => dismiss(id), 3500);
  }, [dismiss]);

  return (
    <ToastContext.Provider value={showToast}>
      {children}
      <div className="toast-container">
        {toasts.map(t => (
          <div key={t.id} className={`toast toast-${t.type}`} onClick={() => dismiss(t.id)} role="status">
            <span aria-hidden="true">{t.type === 'error' ? '⚠' : '✓'}</span>
            <span>{t.message}</span>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}

export function useToast() {
  return useContext(ToastContext);
}

// Small inline spinner for buttons/loading states — pure Tailwind, no extra CSS needed.
export function Spinner() {
  return <span className="inline-block w-4 h-4 border-2 border-current border-t-transparent rounded-full animate-spin" aria-hidden="true" />;
}

const BUTTON_VARIANT_CLASS = {
  primary: 'btn-blue', success: 'btn-green', danger: 'btn-red',
  purple: 'btn-purple', dark: 'btn-dark', ghost: 'btn-ghost',
};

// One button component for every action across all four pages, so "saving" always looks and
// behaves the same way (spinner replaces/joins the label, button disables itself) instead of
// each page hand-rolling its own `{loading ? 'Saving…' : 'Save'}` disable logic.
export function Button({ variant = 'primary', loading = false, disabled = false, children, className = '', ...props }) {
  return (
    <button className={`${BUTTON_VARIANT_CLASS[variant] || BUTTON_VARIANT_CLASS.primary} ${className}`} disabled={disabled || loading} {...props}>
      {loading && <Spinner />}
      {children}
    </button>
  );
}

// Placeholder shape for content that's still loading — communicates "this exact layout is
// coming" instead of a generic "Loading…" line jumping the page around once data arrives.
export function Skeleton({ className = '' }) {
  return <div className={`bg-jira-border rounded animate-pulse ${className}`} aria-hidden="true" />;
}

// A skeleton shaped like PageHeader + a couple of SectionCards — used as the loading state
// for pages that fetch before they can render anything meaningful.
export function PageSkeleton() {
  return (
    <div className="max-w-xl mx-auto p-6" aria-busy="true" aria-label="Loading">
      <div className="flex items-center gap-3 mb-6">
        <Skeleton className="w-8 h-8 rounded" />
        <Skeleton className="h-5 w-48" />
      </div>
      <div className="section-card">
        <Skeleton className="h-4 w-32 mb-3" />
        <Skeleton className="h-8 w-full mb-2" />
        <Skeleton className="h-8 w-full" />
      </div>
      <div className="section-card">
        <Skeleton className="h-4 w-40 mb-3" />
        <Skeleton className="h-8 w-full" />
      </div>
    </div>
  );
}

// Empty/zero-state — an icon + message instead of blank space or a bare "No items." line.
export function EmptyState({ icon = '🗒️', title, description }) {
  return (
    <div className="flex flex-col items-center text-center py-8 px-4">
      <div className="text-2xl mb-2" aria-hidden="true">{icon}</div>
      {title && <div className="text-sm font-semibold text-jira-dark mb-1">{title}</div>}
      {description && <div className="text-xs text-jira-grey max-w-xs">{description}</div>}
    </div>
  );
}

// Forge/Graph failures surface as an Error whose .message is a raw technical dump (a nested
// JSON error body, sometimes wrapped in "There was an error invoking the function - ..."). An
// enterprise app doesn't show that to the user — this pulls out the one line that matters.
export function formatError(err, fallback = 'Something went wrong. Please try again.') {
  const raw = (err && err.message) || '';
  if (!raw) return fallback;
  const jsonMatch = raw.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      const parsed = JSON.parse(jsonMatch[0]);
      const code = parsed.error?.code;
      const message = parsed.error?.message;
      if (code === 'TooManyRequests') return 'Microsoft is rate-limiting requests right now. Wait a few seconds and try again.';
      if (code && message && message !== 'UnknownError') return `${message} (${code})`;
      if (code) return `Request failed: ${code}. Please try again.`;
    } catch { /* not parseable — fall through to the raw message below */ }
  }
  return raw.length > 160 ? `${raw.slice(0, 160)}…` : raw;
}

// A real table for the debug console (was a dark ad-hoc <div> stack) — one row per field,
// grouped under a rowspan'd timestamp per check, errors highlighted red.
export function LogTable({ entries }) {
  if (!entries.length) {
    return <EmptyState icon="🗒️" title="No activity yet" description="Run a check above to see results here." />;
  }
  return (
    <div className="log-table-wrap">
      <table className="log-table">
        <thead>
          <tr><th>Time</th><th>Field</th><th>Value</th></tr>
        </thead>
        <tbody>
          {entries.map((entry, i) => entry.lines.map((line, j) => (
            <tr key={`${i}-${j}`} className={line.isErr ? 'log-row-err' : ''}>
              {j === 0 && <td rowSpan={entry.lines.length} className="log-time">{entry.time}</td>}
              <td className="log-key">{line.key}</td>
              <td className="log-val">{line.val}</td>
            </tr>
          )))}
        </tbody>
      </table>
    </div>
  );
}

export function PageHeader({ title, subtitle }) {
  return (
    <div className="mb-6">
      <div className="flex items-center gap-3 mb-1">
        <div className="w-8 h-8 bg-jira-blue rounded flex items-center justify-center text-white font-bold text-sm">J</div>
        <h1 className="text-lg font-bold text-jira-dark">{title}</h1>
      </div>
      {subtitle && <p className="text-xs text-jira-grey ml-11">{subtitle}</p>}
    </div>
  );
}

// A titled, bordered group — replaces bare "section-title + loose fields" so each
// logical step reads as one visually distinct unit instead of everything running together.
export function SectionCard({ title, description, children, tone = 'default' }) {
  return (
    <div className={`section-card section-card-${tone}`}>
      {title && <div className="section-card-title">{title}</div>}
      {description && <p className="section-card-desc">{description}</p>}
      {children}
    </div>
  );
}

// Horizontal numbered progress indicator — steps before the current one show a
// checkmark (done), the current one is filled and labeled active, the rest are outlined.
export function Stepper({ steps, currentIndex }) {
  return (
    <div className="stepper" role="list" aria-label="Setup progress">
      {steps.map((label, i) => {
        const state = i < currentIndex ? 'done' : i === currentIndex ? 'active' : 'upcoming';
        return (
          <React.Fragment key={label}>
            <div className={`stepper-item stepper-item-${state}`} role="listitem">
              <div className="stepper-dot">{state === 'done' ? '✓' : i + 1}</div>
              <div className="stepper-label">{label}</div>
            </div>
            {i < steps.length - 1 && <div className={`stepper-line ${i < currentIndex ? 'stepper-line-done' : ''}`} />}
          </React.Fragment>
        );
      })}
    </div>
  );
}
