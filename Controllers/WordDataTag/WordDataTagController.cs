using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordDataTag
{
    public class WordDataTagController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //定义WordDocument对象
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //定义DataTag对象
            PageOfficeNetCore.WordWriter.DataTag deptTag = doc.OpenDataTag("{部门名}");
            deptTag.Font.Color = Color.Green;
            //给DataTag对象赋值
            deptTag.Value = "技术";

            PageOfficeNetCore.WordWriter.DataTag userTag = doc.OpenDataTag("{姓名}");
            userTag.Font.Color = Color.Green;
            userTag.Value = "李四";

            PageOfficeNetCore.WordWriter.DataTag dateTag = doc.OpenDataTag("【时间】");
            dateTag.Font.Color = Color.Blue;
            dateTag.Value = DateTime.Now.ToString("yyyy-MM-dd");

            pageofficeCtrl.SetWriter(doc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test2.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}