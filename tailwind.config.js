/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        // NewShire teal — primary brand
        teal: {
          50: '#E8F0F0',
          100: '#D4E5E5',
          400: '#4A8A8B',
          500: '#2C7273',
          700: '#1A5A5B',
          900: '#0F3D3E',
        },
        // NewShire gold — accent
        gold: {
          50: '#FAF5E8',
          200: '#F0E4C0',
          500: '#C9A961',
          700: '#9C7F40',
        },
        // Semantic
        success: '#16A34A',
        warning: '#D97706',
        error: '#DC2626',
        info: '#0284C7',
      },
      fontFamily: {
        sans: ['Source Sans Pro', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'system-ui', 'sans-serif'],
        mono: ['Source Code Pro', 'Consolas', 'Monaco', 'monospace'],
      },
      boxShadow: {
        'drawer': '2px 0 12px rgba(0, 0, 0, 0.08)',
        'card': '0 1px 3px rgba(0, 0, 0, 0.05), 0 1px 2px rgba(0, 0, 0, 0.03)',
      },
      animation: {
        'slide-in': 'slide-in 0.22s ease-out',
        'fade-in': 'fade-in 0.2s ease-out',
      },
      keyframes: {
        'slide-in': {
          from: { transform: 'translateX(-100%)' },
          to: { transform: 'translateX(0)' },
        },
        'fade-in': {
          from: { opacity: '0' },
          to: { opacity: '1' },
        },
      },
    },
  },
  plugins: [],
};
