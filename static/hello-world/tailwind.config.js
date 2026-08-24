/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./src/**/*.{js,jsx,ts,tsx}'],
  theme: {
    extend: {
      colors: {
        // Matches the hardcoded hex values already used via bg-[#...]/border-[#...] elsewhere
        // in these same files (btn-blue, btn-dark, btn-purple, .card, .form-input, etc.) —
        // these tokens were referenced throughout the four page components but never actually
        // defined here, so Tailwind silently generated no CSS for any of them.
        'jira-blue':   '#0052CC',
        'jira-dark':   '#172B4D',
        'jira-grey':   '#6B778C',
        'jira-green':  '#00875A',
        'jira-purple': '#403294',
        'jira-border': '#DFE1E6',
        'jira-light':  '#F4F5F7',
      },
    },
  },
  plugins: [],
};
