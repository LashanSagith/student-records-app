import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base ("./") instead of a hardcoded "/repo-name/" — this makes the
// built asset paths work on GitHub Pages no matter what you name the repo,
// so there's nothing to edit here even if you rename it later.
export default defineConfig({
  plugins: [react()],
  base: "./",
  build: {
    outDir: "dist",
  },
});
