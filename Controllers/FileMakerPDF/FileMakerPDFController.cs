using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.FileMakerPDF
{
    public class FileMakerPDFController : Controller
    {

        private readonly IWebHostEnvironment _webHostEnvironment;

        public FileMakerPDFController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            string url = "";
            url = _webHostEnvironment.WebRootPath;
            ViewBag.url = url + "\\FileMakerPDF\\doc";
            return View();
        }
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
            pdfCtrl.AddCustomToolButton("左转", "RotateLeft()", 12);
            pdfCtrl.AddCustomToolButton("右转", "RotateRight()", 13);
            //pdfCtrl1.AddCustomToolButton("-", "", 0);
            //pdfCtrl1.AddCustomToolButton("放大", "ZoomIn()", 14);
            //pdfCtrl1.AddCustomToolButton("缩小", "ZoomOut()", 15);
            //pdfCtrl1.AddCustomToolButton("-", "", 0);
            //pdfCtrl1.AddCustomToolButton("全屏", "SwitchFullScreen()", 4); 
            //pdfCtrl1.AllowCopy = false;
            pdfCtrl.WebOpen("doc/template.pdf");

            ViewBag.pdfCtrl = pdfCtrl.GetHtmlCode("PDFCtrl1");
            return View();
        }
        public IActionResult FileMakerPDF()
        {
            string type = "";
            type = Request.Query["type"];
            ViewBag.type = type;

            PageOfficeNetCore.FileMakerCtrl fileMakerCtrl = new PageOfficeNetCore.FileMakerCtrl(Request);
            fileMakerCtrl.ServerPage = "/POserver";
            //设置保存页面
            fileMakerCtrl.SaveFilePage = "SaveDoc";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //禁用右击事件
            doc.DisableWindowRightClick = true;
            //给数据区域赋值，即把数据填充到模板中相应的位置
            doc.OpenDataRegion("PO_company").Value = "北京卓正志远软件有限公司";
            fileMakerCtrl.SetWriter(doc);
            fileMakerCtrl.JsFunction_OnProgressComplete = "OnProgressComplete()";
            fileMakerCtrl.FillDocumentAsPDF("../FileMakerPDF/doc/template.doc", PageOfficeNetCore.DocumentOpenType.Word, "a.pdf");
            ViewBag.fmCtrl = fileMakerCtrl.GetHtmlCode("FileMakerCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/FileMakerPDF/doc/" + fs.FileName);
            return  fs.Close();
            
        }
        public ActionResult DownPDF()
        {
            string strFilePath = _webHostEnvironment.WebRootPath + "/FileMakerPDF/doc/" + "template.pdf";//服务器文件路径
            var stream = System.IO.File.OpenRead(strFilePath);
            return File(stream, "application/pdf", "template.pdf");

        }

    }
}