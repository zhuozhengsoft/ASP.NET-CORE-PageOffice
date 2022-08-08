using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DataRegionEdit
{
    public class DataRegionEditController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public DataRegionEditController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            doc.Template.DefineDataRegion("Name", "[ 姓名 ]");
            doc.Template.DefineDataRegion("Address", "[ 地址 ]");
            doc.Template.DefineDataRegion("Tel", "[ 电话 ]");
            doc.Template.DefineDataRegion("Phone", "[ 手机 ]");
            doc.Template.DefineDataRegion("Sex", "[ 性别 ]");
            doc.Template.DefineDataRegion("Age", "[ 年龄 ]");
            doc.Template.DefineDataRegion("Email", "[ 邮箱 ]");
            doc.Template.DefineDataRegion("QQNo", "[ QQ号 ]");
            doc.Template.DefineDataRegion("MSNNo", "[ MSN号 ]");

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("定义数据区域", "ShowDefineDataRegions()", 3);
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
            fs.SaveToFile(webRootPath + "/DataRegionEdit/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}