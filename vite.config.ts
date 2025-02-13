import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue"; // ✅ 确保安装 Vue 插件

export default defineConfig({
  plugins: [vue()], // ✅ 添加 Vue 插件
  optimizeDeps: {
    include: [], // ❌ 移除 "html-docx-js"
  },
  build: {
    commonjsOptions: {
      include: [/node_modules/],
      transformMixedEsModules: true,
    },
  },
});
