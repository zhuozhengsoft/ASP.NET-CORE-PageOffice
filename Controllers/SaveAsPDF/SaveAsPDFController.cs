using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SaveAsPDF
{
    public class SaveAsPDFController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        protected string pdfName = "";
        public SaveAsPDFController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }


        public IActionResult OpenPDF()
        {
            PageOfficeNetCore.PDFCtrl pdfCtrl = new PageOfficeNetCore.PDFCtrl(Request);

            //设置服务器页面
            pdfCtrl.ServerPage = "/POserver";

            pdfCtrl.Theme = PageOfficeNetCore.ThemeType.CustomStyle;
            // 按键说明：光标键、Home、End、PageUp、PageDown可用来移动或翻页；数字键盘+、-用来放大缩小；数字键盘/、*用来旋转页面。
            //AddCustomToolButton方法中的三个参数分别为：按钮名称、按钮执行的JS函数、按钮图标的索引
            pdfCtrl.AddCustomToolButton("打印", "Print()", 6);
            pdfCtrl.AddCustomToolButton("-", "", 0);
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
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("放大", "ZoomIn()", 14);
            pdfCtrl.AddCustomToolButton("缩小", "ZoomOut()", 15);
            pdfCtrl.AddCustomToolButton("-", "", 0);
            pdfCtrl.AddCustomToolButton("全屏", "SwitchFullScreen()", 4);
            pdfCtrl.AllowCopy = false;//是否允许拷贝

            string fileName = Request.Query["fileName"];
            pdfCtrl.WebOpen("doc/" + fileName);

            ViewBag.pdfCtrl = pdfCtrl.GetHtmlCode("PDFCtrl1");
            return View();
        }

        public IActionResult WordToPDF()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("另存为PDF文件", "SaveAsPDF()", 1);
            string fileName = "template.doc";
            //定义将要转换的PDF文件的名称
            pdfName = fileName.Substring(0, fileName.Length - 4) + ".pdf";

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/" + fileName, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.pdfName = pdfName;
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SaveAsPDF/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}