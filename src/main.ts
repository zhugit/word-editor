import { createApp } from "vue";
import App from "./App.vue";
import "blob-polyfill"; // 解决 Blob 兼容问题

const app = createApp(App);
app.mount("#app");
