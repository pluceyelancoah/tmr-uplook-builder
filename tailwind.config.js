/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./index.html", "./js/**/*.js"],
  theme: { extend: {
    fontFamily: { sans: ['Inter','ui-sans-serif','system-ui','Segoe UI','Arial','sans-serif'] },
    colors: {
      border:'hsl(214 25% 90%)', input:'hsl(214 25% 90%)', ring:'hsl(212 70% 38%)',
      background:'hsl(210 30% 98%)', foreground:'hsl(215 28% 17%)',
      muted:'hsl(214 25% 96%)', 'muted-foreground':'hsl(215 16% 47%)',
      card:'hsl(0 0% 100%)',
      primary:'hsl(212 64% 30%)', 'primary-foreground':'hsl(0 0% 100%)',
      brand:'hsl(212 64% 30%)',
      success:'hsl(142 55% 30%)', 'success-bg':'hsl(140 55% 96%)', 'success-border':'hsl(142 45% 80%)',
      danger:'hsl(0 65% 42%)', 'danger-bg':'hsl(0 75% 97%)', 'danger-border':'hsl(0 70% 88%)',
    },
    borderRadius: { lg:'0.6rem', xl:'0.85rem' },
    boxShadow: { card:'0 1px 2px 0 rgb(16 24 40 / 0.04), 0 1px 3px 0 rgb(16 24 40 / 0.06)' },
  }},
};
