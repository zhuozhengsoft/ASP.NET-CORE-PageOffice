using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DataTagEdit
{
    public class DataTagEditController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public DataTagEditController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            doc.Template.DefineDataTag("{ 甲方 }");
            doc.Template.DefineDataTag("{ 乙方 }");
            doc.Template.DefineDataTag("{ 担保人 }");
            doc.Template.DefineDataTag("【 合同日期 】");
            doc.Template.DefineDataTag("【 合同编号 】");

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("定义数据标签", "ShowDefineDataTags()", 20);
            pageofficeCtrl.Theme = PageOfficeNetCore.ThemeType.Office2007;
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            pageofficeCtrl.SetWriter(doc);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/DataTagEdit/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}