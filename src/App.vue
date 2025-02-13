<template>
  <div class="app-container">
    <h1 class="app-title">Vue-word在线编辑器</h1>
    <WordEditor />
  </div>
</template>

<script>
import WordEditor from "./components/WordEditor.vue";

export default {
  components: { WordEditor },
};
</script>

<style>
.app-container {
  text-align: center; /* 使容器内的内容居中 */
  padding: 20px;
}

.app-title {
  font-size: 36px;
  font-weight: bold;
  color: #333;
  margin: 0;
  margin-bottom: 30px; /* 添加标题与编辑器之间的间距 */
}
</style>
