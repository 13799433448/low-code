import { write, utils } from "xlsx";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
// Hook 导出html文件
import exportHtmlHook from "@/hook/exportHtml";

export default function () {
  // 导出excel
  function exportExcel(Dom: any) {
    let et = utils.table_to_book(Dom); //此处传入table的DOM节点
    let etout = write(et, {
      bookType: "xlsx",
      bookSST: true,
      type: "array",
    });
    try {
      downFileToLocal("自定义表格.xlsx", new Blob([etout], {
        type: "application/octet-stream",
      }),);
    } catch (e) {
      console.log(e, etout);
    }
    return etout;
  }
  // 导出图片
  function onSaveCanvas(table: HTMLElement) {
    // 这里的类名要与点击事件里的一样
    html2canvas(table, { scale: 2, logging: false, useCORS: true }).then(function (canvas: {
      toDataURL: (arg0: string) => any;
    }) {
      const type = "png";
      let imgData = canvas.toDataURL(type);
      // 图片格式处理
      let _fixType = function (type: any) {
        type = type.toLowerCase().replace(/jpg/i, "jpeg");
        let r = type.match(/png|jpeg|bmp|gif/)[0];
        return "image/" + r;
      };
      imgData = imgData.replace(_fixType(type), "image/octet-stream");
      let filename = "htmlImg" + "." + type;
      // 保存为文件
      //  以bolb文件下载
      downFileToLocal(filename, convertBase64ToBlob(imgData));
    });
  }
  // 导出为PDF
  function onPDFExport(table: HTMLElement) {
    html2canvas(table).then(function (canvas) {
      let contentWidth = canvas.width;
      let contentHeight = canvas.height;
      //一页pdf显示html页面生成的canvas高度;
      let pageHeight = (contentWidth / 592.28) * 841.89;
      //未生成pdf的html页面高度
      let leftHeight = contentHeight;
      //页面偏移
      let position = 0;
      //a4纸的尺寸[595.28,841.89]，html页面生成的canvas在pdf中图片的宽高
      let imgWidth = 595.28;
      let imgHeight = (592.28 / contentWidth) * contentHeight;

      let pageData = canvas.toDataURL("image/jpeg", 1.0);

      let pdf = new jsPDF("p", "pt", "a4");

      //有两个高度需要区分，一个是html页面的实际高度，和生成pdf的页面高度(841.89)
      //当内容未超过pdf一页显示的范围，无需分页
      if (leftHeight < pageHeight) {
        pdf.addImage(pageData, "JPEG", 0, 0, imgWidth, imgHeight);
      } else {
        while (leftHeight > 0) {
          pdf.addImage(pageData, "JPEG", 0, position, imgWidth, imgHeight);
          leftHeight -= pageHeight;
          position -= 841.89;
          //避免添加空白页
          if (leftHeight > 0) {
            pdf.addPage();
          }
        }
      }
      pdf.save("content.pdf");
    });
  }

  // 导出html
  function onExportHtml (table: any) {
    let template = table.outerHTML;
    downFileToLocal(
      "模板.html",
      new Blob([exportHtmlHook().getHtml(template)], {
        type: "",
      })
    );
  };
  // base64转化为Blob对象
  function convertBase64ToBlob(imageEditorBase64: string) {
    let base64Arr = imageEditorBase64.split(",");
    let imgtype = "";
    let base64String = "";
    if (base64Arr.length > 1) {
      //如果是图片base64，去掉头信息
      base64String = base64Arr[1];
      imgtype = base64Arr[0].substring(base64Arr[0].indexOf(":") + 1, base64Arr[0].indexOf(";"));
    }
    // 将base64解码
    let bytes = atob(base64String);
    //let bytes = base64;
    let bytesCode = new ArrayBuffer(bytes.length);
    // 转换为类型化数组
    let byteArray = new Uint8Array(bytesCode);

    // 将base64转换为ascii码
    for (let i = 0; i < bytes.length; i++) {
      byteArray[i] = bytes.charCodeAt(i);
    }
    // 生成Blob对象（文件对象）
    return new Blob([bytesCode], { type: imgtype });
  }
  // 下载Blob流文件
  function downFileToLocal(fileName: string, blob: Blob) {
    // 创建用于下载文件的a标签
    const d = document.createElement("a");
    // 设置下载内容
    d.href = URL.createObjectURL(blob);
    // 设置下载文件的名字
    d.download = fileName;
    // 界面上隐藏该按钮
    d.style.display = "none";
    // 放到页面上
    document.body.appendChild(d);
    // 点击下载文件
    d.click();
    // 从页面移除掉
    document.body.removeChild(d);
    // 释放 URL.createObjectURL() 创建的 URL 对象
    window.URL.revokeObjectURL(d.href);
  }
  return {
    exportExcel,
    onSaveCanvas,
    onPDFExport,
    onExportHtml
  };
}
