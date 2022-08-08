using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.POPDF2
{
    public class POPDF2Controller : Controller
    {
        public IActionResult PDF()
        {
            // 按键说明：光标键、Home、End、PageUp、PageDown可用来移动或翻页；数字键盘+、-用来放大缩小；数字键盘/、*用来旋转页面。

            PageOfficeNetCore.PDFCtrl pdfCtrl = new PageOfficeNetCore.PDFCtrl(Request);

            pdfCtrl.ServerPage = "/POserver";

            pdfCtrl.Theme = PageOfficeNetCore.ThemeType.Office2007;
            //pdfCtrl1.TitlebarColor = Color.Green;
            //pdfCtrl1.JsFunction_AfterDocumentOpened = "AfterDocumentOpened()";
            pdfCtrl.AddCustomToolButton("打印", "Print()", 6);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("显示/隐藏书签", "SwitchBKMK()", 0);
            pdfCtrl.AddCustomToolButton("实际大小", "SetPageReal()", 16);
            pdfCtrl.AddCustomToolButton("适合页面", "SetPageFit()", 17);
            pdfCtrl.AddCustomToolButton("适合宽度", "SetPageWidth()", 18);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("首页", "FirstPage()", 8);
            pdfCtrl.AddCustomToolButton("上一页", "PreviousPage()", 9);
            pdfCtrl.AddCustomToolButton("下一页", "NextPage()", 10);
            pdfCtrl.AddCustomToolButton("尾页", "LastPage()", 11);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            //pdfCtrl1.AddCustomToolButton("左转", "RotateLeft()", 12);
            //pdfCtrl1.AddCustomToolButton("右转", "RotateRight()", 13);
            //pdfCtrl1.AddCustomToolButton("-", "", 0);
            //pdfCtrl1.AddCustomToolButton("放大", "ZoomIn()", 14);
            //pdfCtrl1.AddCustomToolButton("缩小", "ZoomOut()", 15);
            //pdfCtrl1.AddCustomToolButton("-", "", 0);
            //pdfCtrl1.AddCustomToolButton("全屏", "SwitchFullScreen()", 4); 
            //pdfCtrl1.AllowCopy = false;
            pdfCtrl.WebOpen("doc/test2.pdf");

            ViewBag.pdfCtrl = pdfCtrl.GetHtmlCode("PDFCtrl1");
            return View();
        }
    }
}