using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.PDFSearch
{
    public class PDFSearchController : Controller
    {
        public IActionResult PDF()
        {
            PageOfficeNetCore.PDFCtrl pdfCtrl = new PageOfficeNetCore.PDFCtrl(Request);
            pdfCtrl.ServerPage = "/POserver";

            pdfCtrl.Theme = PageOfficeNetCore.ThemeType.Office2007;
            pdfCtrl.AddCustomToolButton("搜索", "SearchText()", 0);
            pdfCtrl.AddCustomToolButton("搜索下一个", "SearchTextNext()", 0);
            pdfCtrl.AddCustomToolButton("搜索上一个", "SearchTextPrev()", 0);
            pdfCtrl.AddCustomToolButton("实际大小", "SetPageReal()", 16);
            pdfCtrl.AddCustomToolButton("适合页面", "SetPageFit()", 17);
            pdfCtrl.AddCustomToolButton("适合宽度", "SetPageWidth()", 18);
            //打开Word文档
            pdfCtrl.WebOpen("doc/test.pdf");
            ViewBag.POCtrl = pdfCtrl.GetHtmlCode("PDFCtrl1");
            return View();
        }
    }
}