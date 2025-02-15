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
import { Document, Packer, Paragraph, TextRun, ImageRun, ExternalHyperlink } from "docx";
import { jsPDF } from "jspdf";
import mammoth from "mammoth";

// **加载字体**
const loadFontAsBase64 = async () => {
  const fontUrl = "/fonts/NotoSerifCJKsc-VF.ttf"; // 确保路径正确
  const response = await fetch(fontUrl);
  const buffer = await response.arrayBuffer();
  return btoa(
    new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), "")
  );
};

export default {
  components: {
    QuillEditor,
  },
  setup() {
    const editorContent = ref("");

    // **自定义工具栏**
    const editorOptions = {
      modules: {
        toolbar: [
          [{ font: [] }, { size: [] }],
          ["bold", "italic", "underline", "strike"],
          [{ color: [] }, { background: [] }],
          [{ script: "sub" }, { script: "super" }],
          [{ align: [] }],
          [{ list: "ordered" }, { list: "bullet" }],
          [{ indent: "-1" }, { indent: "+1" }],
          ["link", "image", "video"],
          ["clean"],
        ],
      },
      theme: "snow",
    };

    // **导出 Word**
    const exportWord = async () => {
      let filename = prompt("请输入文件名", "文档");
      if (!filename) return;

      const htmlContent = editorContent.value;
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = htmlContent;

      const docChildren = [];

      for (const child of tempDiv.childNodes) {
        if (child.nodeType === 3 && child.textContent.trim()) {
          // 处理普通文本
          docChildren.push(new Paragraph({ children: [new TextRun(child.textContent)] }));
        } else if (child.nodeType === 1 && child.tagName === "P") {
          // 处理段落
          const paragraphChildren = [];

          for (const innerChild of child.childNodes) {
            if (innerChild.nodeType === 3 && innerChild.textContent.trim()) {
              paragraphChildren.push(new TextRun(innerChild.textContent));
            } else if (innerChild.nodeType === 1 && innerChild.tagName === "A") {
              // 可点击的超链接
              const url = innerChild.getAttribute("href");
              const linkText = innerChild.textContent || url;
              paragraphChildren.push(
                new ExternalHyperlink({
                  children: [
                    new TextRun({
                      text: linkText,
                      style: "Hyperlink",
                      color: "0000FF", // 超链接蓝色
                      underline: true,
                    }),
                  ],
                  link: url,
                })
              );
            } else if (innerChild.nodeType === 1 && innerChild.tagName === "IMG") {
              const imgSrc = innerChild.getAttribute("src");
              if (imgSrc.startsWith("data:image")) {
                const base64Data = imgSrc.split(",")[1];
                const buffer = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
                // 获取原始尺寸
                const imgWidth = innerChild.width || 600;
                const imgHeight = innerChild.height || (imgWidth * 0.75);
                // A4 适配
                const maxWidth = 450; // A4 适应的最大宽度 (单位: pt)
                const scaleFactor = Math.min(1, maxWidth / imgWidth);

                paragraphChildren.push(
                  new ImageRun({
                    data: buffer,
                    transformation: {
                      width: imgWidth * scaleFactor,
                      height: imgHeight * scaleFactor,
                    },
                  })
                );
              }
            }
          }

          if (paragraphChildren.length > 0) {
            docChildren.push(new Paragraph({ children: paragraphChildren }));
          }
        }
      }
      // 设置 A4 竖版
      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                size: { width: 11907, height: 16840 },
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
              },
            },
            children: docChildren,
          },
        ],
      });

      Packer.toBlob(doc).then((blob) => {
        saveAs(blob, `${filename}.docx`);
      });
    };

    // **导出 PDF**
    const exportPDF = async () => {
      let filename = prompt("请输入文件名", "文档");
      if (!filename) return;

      const pdf = new jsPDF();
      const fontBase64 = await loadFontAsBase64();

      pdf.addFileToVFS("NotoSerifCJKsc.ttf", fontBase64);
      pdf.addFont("NotoSerifCJKsc.ttf", "NotoSerif", "normal");
      pdf.setFont("NotoSerif");

      let y = 10;

      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = editorContent.value;

      for (const child of tempDiv.childNodes) {
        if (child.nodeType === 3 && child.textContent.trim()) {
          pdf.text(child.textContent, 10, y);
          y += 10;
        } else if (child.nodeType === 1 && child.tagName === "P") {
          for (const innerChild of child.childNodes) {
            if (innerChild.nodeType === 3 && innerChild.textContent.trim()) {
              pdf.text(innerChild.textContent, 10, y);
              y += 10;
            } else if (innerChild.nodeType === 1 && innerChild.tagName === "A") {
              const url = innerChild.getAttribute("href");
              pdf.textWithLink(innerChild.textContent, 10, y, { url });
              y += 10;
            } else if (innerChild.nodeType === 1 && innerChild.tagName === "IMG") {
              const imgSrc = innerChild.getAttribute("src");
              if (imgSrc.startsWith("data:image")) {
                const imgData = imgSrc.split(",")[1];
                pdf.addImage(imgData, "PNG", 10, y, 100, 100);
                y += 110;
              }
            }
          }
        }
      }

      pdf.save(`${filename}.pdf`);
    };

    // **导入 Word**
    const importWord = async (event) => {
      const file = event.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = async (e) => {
          try {
            const result = await mammoth.convertToHtml({ arrayBuffer: e.target.result });
            let htmlContent = result.value;
            // 解析 HTML，找到所有 <p> 标签，确保正确插入图片
            const tempDiv = document.createElement("div");
            tempDiv.innerHTML = htmlContent;

            for (const p of tempDiv.getElementsByTagName("p")) {
              if (p.innerHTML.trim() === "") {
                // 如果 <p> 为空，插入占位符以防止空行
                p.innerHTML = "<br>";
              }
            }

            editorContent.value = tempDiv.innerHTML;
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