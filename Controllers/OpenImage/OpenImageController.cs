using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.OpenImage
{
    public class OpenImageController : Controller
    {
        public IActionResult Image()
        {
            PageOfficeNetCore.PDFCtrl pdfCtrl = new PageOfficeNetCore.PDFCtrl(Request);
            pdfCtrl.ServerPage = "/POserver";

            // 按键说明：光标键、Home、End、PageUp、PageDown可用来移动或翻页；数字键盘+、-用来放大缩小；数字键盘/、*用来旋转页面。

            pdfCtrl.Theme = PageOfficeNetCore.ThemeType.Office2007;
            //PDFCtrl1.TitlebarColor = Color.Green;
            //PDFCtrl1.JsFunction_AfterDocumentOpened = "AfterDocumentOpened()";
            pdfCtrl.AddCustomToolButton("打印", "Print()", 6);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("实际大小", "SetPageReal()", 16);
            pdfCtrl.AddCustomToolButton("适合页面", "SetPageFit()", 17);
            pdfCtrl.AddCustomToolButton("适合宽度", "SetPageWidth()", 18);
            pdfCtrl.AddCustomToolButton("-", "", 0);

            pdfCtrl.AddCustomToolButton("左转", "RotateLeft()", 12);
            pdfCtrl.AddCustomToolButton("右转", "RotateRight()", 13);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("放大", "ZoomIn()", 14);
            pdfCtrl.AddCustomToolButton("缩小", "ZoomOut()", 15);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("全屏", "SwitchFullScreen()", 4);
            pdfCtrl.WebOpen("doc/test.jpg");

            ViewBag.POCtrl = pdfCtrl.GetHtmlCode("PDFCtrl1");
            return View();
        }
    }
}