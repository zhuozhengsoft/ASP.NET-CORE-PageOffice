using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.FileMaker
{
    public class FileMakerController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public FileMakerController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            string url = "";
            url = _webHostEnvironment.WebRootPath;
            ViewBag.url = url + "\\FileMaker\\doc";
            return View();
        }
        public IActionResult FileMaker()
        {
            string id = Request.Query["id"];
            PageOfficeNetCore.FileMakerCtrl fileMakerCtrl = new PageOfficeNetCore.FileMakerCtrl(Request);
            fileMakerCtrl.ServerPage = "/POserver";
            //设置保存页面
            fileMakerCtrl.SaveFilePage = "SaveDoc?id=" + id; ;
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //禁用右击事件
            //禁用右击事件
            doc.DisableWindowRightClick = true;
            //给数据区域赋值，即把数据填充到模板中相应的位置
            doc.OpenDataRegion("PO_company").Value = "北京卓正志远软件有限公司  " + id;
            fileMakerCtrl.SetWriter(doc);
            //设置转换完成后执行的JS函数
            fileMakerCtrl.JsFunction_OnProgressComplete = "OnProgressComplete()";
            //打开文档
            fileMakerCtrl.FillDocument("../FileMaker/doc/template.doc", PageOfficeNetCore.DocumentOpenType.Word);

            ViewBag.fmCtrl = fileMakerCtrl.GetHtmlCode("FileMakerCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            string id = Request.Query["id"];
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;

            string fileName = "maker" + id + fs.FileExtName; ;
            fs.SaveToFile(webRootPath + "/FileMaker/doc/" + fileName);
            return  fs.Close();
            
        }
    }
}