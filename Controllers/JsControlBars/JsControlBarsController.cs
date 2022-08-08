using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.JsControlBars
{
    public class JsControlBarsController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            // 设置保存文件页面
            pageofficeCtrl.SaveFilePage = "SaveFile";

            // 添加一个自定义工具条上的按钮，AddCustomToolButton的参数说明，详见开发帮助。
            pageofficeCtrl.AddCustomToolButton("保存", "mySave()", 1);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}