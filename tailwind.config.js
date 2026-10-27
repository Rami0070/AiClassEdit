/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./src/pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/components/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/app/**/*.{js,ts,jsx,tsx,mdx}",
    "./index.html"
  ],
  theme: {
    extend: {
      colors: {
        mint: {
          50: '#F0FDF9',
          100: '#CCFBF1',
          200: '#99F6E4',
          300: '#5EEAD4',
          400: '#2DD4BF',
          500: '#14B8A6',
          600: '#0D9488',
          700: '#0F766E',
          800: '#115E59',
          900: '#134E4A',
        },
        sage: {
          50: '#F4F7F5',
          100: '#E4EBE6',
          200: '#CCD8D0',
          300: '#B0C2B5',
          400: '#8FA896',
          500: '#718E7A',
          600: '#58715F',
          700: '#46594C',
        }
      }
    },
  },
  plugins: [],
};
