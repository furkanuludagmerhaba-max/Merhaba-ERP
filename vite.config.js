import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "/Merhaba-ERP/",
  plugins: [react()],
  build: { chunkSizeWarningLimit: 2000 },
});
