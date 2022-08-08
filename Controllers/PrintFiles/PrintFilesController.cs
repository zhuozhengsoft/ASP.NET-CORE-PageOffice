using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.PrintFiles
{
    public class PrintFilesController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public PrintFilesController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Index()
        {
            string url = _webHostEnvironment.WebRootPath + " /PrintFiles/doc/";
            ViewBag.url = url;
            return View();
        }

        public IActionResult Print()
        {

            PageOfficeNetCore.FileMakerCtrl fileMakerCtrl = new PageOfficeNetCore.FileMakerCtrl(Request);
            fileMakerCtrl.ServerPage = "/POserver";

            string id = Request.Query["id"];
            if (id != null && id.Length > 0)
            {
                PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
                //禁用右击事件
                doc.DisableWindowRightClick = true;
                //给数据区域赋值，即把数据填充到模板中相应的位置
                doc.OpenDataRegion("PO_company").Value = "北京卓正志远软件有限公司  " + id;
                //设置保存页面
                fileMakerCtrl.SaveFilePage = "SaveDoc?id=" + id;
                fileMakerCtrl.SetWriter(doc);
                //设置转换完成后执行的JS函数
                fileMakerCtrl.JsFunction_OnProgressComplete = "OnProgressComplete()";
                //打开文档
                fileMakerCtrl.FillDocument("../PrintFiles/doc/template.doc", PageOfficeNetCore.DocumentOpenType.Word);
            }

            ViewBag.fmCtrl = fileMakerCtrl.GetHtmlCode("FileMakerCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {

            string id = Request.Query["id"];

            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string fileName = "maker" + id + fs.FileExtName;
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/PrintFiles/doc/" + fileName);
            return fs.Close();
            
        }

    }
}