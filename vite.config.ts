import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { resolve } from "node:path";

// tauri dev expects the vite server on a fixed port and never auto-opens it.
// dev import-analysis does not follow tsconfig "paths", so alias @/ by hand
export default defineConfig({
  plugins: [react(), tailwindcss()],
  clearScreen: false,
  resolve: { alias: { "@": resolve(import.meta.dirname, "src") } },
  server: { port: 1420, strictPort: true },
  build: {
    rollupOptions: {
      input: {
        main: resolve(import.meta.dirname, "index.html"),
        preview: resolve(import.meta.dirname, "preview.html"),
      },
    },
  },
});
