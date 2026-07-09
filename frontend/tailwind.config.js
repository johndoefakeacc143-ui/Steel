/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        steel: {
          50: "#f4f7f8",
          100: "#e3ebef",
          200: "#c9d8e0",
          300: "#9fb8c6",
          400: "#6e92a6",
          500: "#527589",
          600: "#455f72",
          700: "#3b4f5e",
          800: "#344350",
          900: "#2e3944",
          950: "#1a2229",
        },
        ember: {
          400: "#e08a3c",
          500: "#c96a1a",
          600: "#a85412",
        },
      },
      fontFamily: {
        display: ['"Barlow Condensed"', "sans-serif"],
        sans: ['"Source Sans 3"', "sans-serif"],
      },
      backgroundImage: {
        grid: "linear-gradient(rgba(46,57,68,0.06) 1px, transparent 1px), linear-gradient(90deg, rgba(46,57,68,0.06) 1px, transparent 1px)",
        forge: "radial-gradient(ellipse 80% 60% at 20% 10%, rgba(201,106,26,0.18), transparent 55%), radial-gradient(ellipse 70% 50% at 90% 20%, rgba(82,117,137,0.22), transparent 50%), linear-gradient(165deg, #f4f7f8 0%, #e3ebef 45%, #d5e0e6 100%)",
      },
      backgroundSize: {
        grid: "32px 32px",
      },
      keyframes: {
        rise: {
          "0%": { opacity: "0", transform: "translateY(16px)" },
          "100%": { opacity: "1", transform: "translateY(0)" },
        },
        pulseBar: {
          "0%, 100%": { opacity: "0.55" },
          "50%": { opacity: "1" },
        },
        scan: {
          "0%": { transform: "translateY(-100%)" },
          "100%": { transform: "translateY(320%)" },
        },
      },
      animation: {
        rise: "rise 0.55s ease-out both",
        "pulse-bar": "pulseBar 1.6s ease-in-out infinite",
        scan: "scan 2.4s ease-in-out infinite",
      },
    },
  },
  plugins: [],
};
