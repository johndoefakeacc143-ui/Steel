/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        steel: {
          50: "#f3f6f8",
          100: "#e2e9ee",
          200: "#c5d3dd",
          300: "#9bb2c3",
          400: "#6b8ba3",
          500: "#4f7088",
          600: "#3f5a6f",
          700: "#354a5b",
          800: "#2f3f4d",
          900: "#2a3642",
          950: "#1a2229",
        },
        ember: {
          400: "#d9773a",
          500: "#c45f22",
          600: "#a34b1a",
        },
      },
      fontFamily: {
        display: ['"Bebas Neue"', "Impact", "sans-serif"],
        sans: ['"Source Sans 3"', "Segoe UI", "sans-serif"],
        mono: ['"IBM Plex Mono"', "ui-monospace", "monospace"],
      },
      backgroundImage: {
        blueprint:
          "linear-gradient(rgba(42,54,66,0.04) 1px, transparent 1px), linear-gradient(90deg, rgba(42,54,66,0.04) 1px, transparent 1px)",
      },
      backgroundSize: {
        blueprint: "28px 28px",
      },
      keyframes: {
        "fade-up": {
          "0%": { opacity: "0", transform: "translateY(16px)" },
          "100%": { opacity: "1", transform: "translateY(0)" },
        },
        "scan-line": {
          "0%": { transform: "translateY(-100%)" },
          "100%": { transform: "translateY(400%)" },
        },
        pulseSoft: {
          "0%, 100%": { opacity: "0.55" },
          "50%": { opacity: "1" },
        },
      },
      animation: {
        "fade-up": "fade-up 0.55s ease-out both",
        "scan-line": "scan-line 2.4s ease-in-out infinite",
        pulseSoft: "pulseSoft 1.8s ease-in-out infinite",
      },
    },
  },
  plugins: [],
};
