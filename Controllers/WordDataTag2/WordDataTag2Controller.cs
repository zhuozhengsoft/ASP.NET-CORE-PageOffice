using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordDataTag2
{
    public class WordDataTag2Controller : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //定义WordDocument对象
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //定义DataTag对象
            PageOfficeNetCore.WordWriter.DataTag deptTag = doc.OpenDataTag("{部门名}");
            deptTag.Value = "技术";

            PageOfficeNetCore.WordWriter.DataTag userTag = doc.OpenDataTag("{姓名}");
            userTag.Value = "李志";

            PageOfficeNetCore.WordWriter.DataTag dateTag = doc.OpenDataTag("【时间】");
            dateTag.Value = DateTime.Now.ToString("yyyy-MM-dd");

            pageofficeCtrl.SetWriter(doc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test2.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}