<template>
    <div class="editor-container">
  
      <!-- 富文本编辑器 -->
      <quill-editor
        ref="quillEditor"
        v-model:content="editorContent"
        content-type="html"
        theme="snow"
        :options="editorOptions"
        class="quill-editor"
      />
  
      <!-- 导出和导入功能 -->
      <div class="toolbar">
        <button @click="exportWord">📄 导出 Word</button>
        <button @click="exportPDF">📄 导出 PDF</button>
        <input type="file" @change="importWord" accept=".docx" />
      </div>
    </div>
  </template>
  
  <script>
  import { ref } from "vue";
  import { QuillEditor } from "@vueup/vue-quill";
  import "quill/dist/quill.snow.css";
  import { saveAs } from "file-saver";
  import { Document, Packer, Paragraph, TextRun } from "docx";
  import mammoth from "mammoth";
  import { jsPDF } from "jspdf";
  
  export default {
    components: {
      QuillEditor,
    },
    setup() {
      const editorContent = ref("");
  
      // 自定义工具栏，尽量接近 WPS
      const editorOptions = {
        modules: {
          toolbar: [
            [{ font: [] }, { size: [] }], // 字体 & 字号
            ["bold", "italic", "underline", "strike"], // 加粗, 斜体, 下划线, 删除线
            [{ color: [] }, { background: [] }], // 字体颜色 & 背景色
            [{ script: "sub" }, { script: "super" }], // 上标 & 下标
            [{ align: [] }], // 对齐方式
            [{ list: "ordered" }, { list: "bullet" }], // 列表
            [{ indent: "-1" }, { indent: "+1" }], // 缩进
            ["link", "image", "video"], // 插入超链接, 图片, 视频
            ["clean"], // 清除格式
          ],
        },
        theme: "snow",
      };
  
      // 导出 Word（保持格式）
      const exportWord = () => {
        let filename = prompt("请输入文件名", "文档");
        if (!filename) return;
  
        const doc = new Document({
          sections: [
            {
              children: [
                new Paragraph({
                  children: [new TextRun(editorContent.value)],
                }),
              ],
            },
          ],
        });
  
        Packer.toBlob(doc).then((blob) => {
          saveAs(blob, filename + ".docx");
        });
      };
  
      // 导出 PDF
      const exportPDF = () => {
        let filename = prompt("请输入文件名", "文档");
        if (!filename) return;
  
        const doc = new jsPDF();
        doc.text(editorContent.value, 10, 10);
        doc.save(filename + ".pdf");
      };
  
      // 导入 Word（解析 .docx 文件）
      const importWord = (event) => {
        const file = event.target.files[0];
        if (file) {
          const reader = new FileReader();
          reader.onload = async (e) => {
            try {
              const result = await mammoth.convertToHtml({ arrayBuffer: e.target.result });
              editorContent.value = result.value;
            } catch (error) {
              console.error("Word 文件解析失败：", error);
            }
          };
          reader.readAsArrayBuffer(file);
        }
      };
  
      return {
        editorContent,
        editorOptions,
        exportWord,
        exportPDF,
        importWord,
      };
    },
  };
  </script>
  
  <style>
  .editor-container {
    border: 1px solid #ddd;
    border-radius: 5px;
    width: 80%;
    margin: auto;
    background: #fff;
    padding: 10px;
  }
  .toolbar {
    padding: 10px;
    background: #f5f5f5;
    text-align: center;
  }
  .toolbar button {
    margin: 5px;
    padding: 5px 10px;
    cursor: pointer;
  }
  .quill-editor {
    height: 400px;
  }
  </style>
  