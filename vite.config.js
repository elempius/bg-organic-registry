import { defineConfig } from "vite";

export default defineConfig({
  // Relative base so the build works at any GitHub Pages path
  // (project pages /repo/ or a custom domain at root).
  base: "./",
  build: {
    target: "es2020",
    chunkSizeWarningLimit: 1500,
  },
});
