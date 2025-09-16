import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "path";
import tailwindcss from "@tailwindcss/vite";

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    outDir: "../dist",
  },
  resolve: {
    alias: {
      "@": resolve(__dirname, "src"),
    },
  },
  server: {
    watch: {
      usePolling: true,
    },
  },
});
