import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "path";

console.log("Vite root is:", new URL(".", import.meta.url).pathname);

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "../dist",
  },
  resolve: {
    alias: {
      "@": resolve(__dirname, "client/src"),
    },
  },
  server: {
    watch: {
      usePolling: true,
    },
  },
});
