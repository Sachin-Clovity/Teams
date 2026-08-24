import React from 'react';

// ── Shared page chrome — used by all four Jira-embedded pages so they read as one
// consistent app instead of four independently-built forms. ─────────────────────

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
