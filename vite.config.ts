import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      events: "events/",
      util: "util/",
      stream: "stream-browserify/",
    },
  },
  optimizeDeps: {
    include: ["docx", "xmlbuilder2", "html-to-docx", "events", "util", "stream-browserify"],
  },
  define: {
    global: {},
    process: {
      env: {},
    },
  },
  build: {
    commonjsOptions: {
      include: [/node_modules/],
      transformMixedEsModules: true,
    },
  },
});
