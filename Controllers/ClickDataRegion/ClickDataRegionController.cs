using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ClickDataRegion
{
    public class ClickDataRegionController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public ClickDataRegionController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion dataReg = doc.OpenDataRegion("PO_deptName");
            //为方便用户知道哪些地方可以编辑，所以设置了数据区域的背景色
            dataReg.Shading.BackgroundPatternColor = Color.Pink;
            //dataReg.Editing = true;

            pageofficeCtrl.SetWriter(doc); // 不要忘记此句代码

            // 设置数据区域点击时的响应js函数
            pageofficeCtrl.JsFunction_OnWordDataRegionClick = "OnWordDataRegionClick()";

            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.OfficeToolbars = false;
            pageofficeCtrl.Caption = "为方便用户知道哪些地方可以编辑，所以设置了数据区域的背景色";
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/ClickDataRegion/doc/" + fs.FileName);
            return  fs.Close();
            
        }
    }
}