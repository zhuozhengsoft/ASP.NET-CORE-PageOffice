using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.FileMakerConvertPDFs
{
    public class FileMakerConvertPDFsController : Controller
    {

        private readonly IWebHostEnvironment _webHostEnvironment;

        public FileMakerConvertPDFsController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Index()
        {
            string url = "";
            url = _webHostEnvironment.WebRootPath;
            ViewBag.url = url + "\\FileMakerConvertPDFs\\doc";
            return View();
        }
        public IActionResult Word()
        {
            String id = Request.Query["id"];
            String filePath = "";
            string webRootPath = _webHostEnvironment.WebRootPath;
            if ("1".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/PageOffice产品简介.doc";
            }
            if ("2".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/Pageoffice客户端安装步骤.doc";

            }
            if ("3".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/PageOffice的应用领域.doc";
            }
            if ("4".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/PageOffice产品对客户端环境要求.doc";
            }

            filePath = filePath.Replace("/", "\\");

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen(filePath, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult Convert()
        {
            String id = Request.Query["id"];
            String filePath = "";
            string webRootPath = _webHostEnvironment.WebRootPath;
            if ("1".Equals(id))
            {
                //filePath = webRootPath+"/FileMakerConvertPDFs/doc/PageOffice产品简介.doc";
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/PageOffice产品简介.doc";
            }
            if ("2".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/Pageoffice客户端安装步骤.doc";

            }
            if ("3".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/PageOffice的应用领域.doc";
            }
            if ("4".Equals(id))
            {
                filePath = webRootPath + "/FileMakerConvertPDFs/doc/PageOffice产品对客户端环境要求.doc";
            }
            filePath = filePath.Replace("/", "\\");

            PageOfficeNetCore.FileMakerCtrl fileMakerCtrl = new PageOfficeNetCore.FileMakerCtrl(Request);
            fileMakerCtrl.ServerPage = "/POserver";
            //设置保存页面
            fileMakerCtrl.SaveFilePage = "SaveDoc";
            //设置转换完成后执行的JS函数
            fileMakerCtrl.JsFunction_OnProgressComplete = "OnProgressComplete()";
            //打开文档
            fileMakerCtrl.FillDocumentAsPDF(filePath, PageOfficeNetCore.DocumentOpenType.Word, "aa.pdf");
            ViewBag.fmCtrl = fileMakerCtrl.GetHtmlCode("FileMakerCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/FileMakerConvertPDFs/doc/" + fs.FileName);
            return  fs.Close();
            
        }

    }
}