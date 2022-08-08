using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.InsertImageSetSize
{
    public class InsertImageSetSizeController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion data1 = doc.OpenDataRegion("PO_p1");
            data1.Value = "[image width=200.2 height=200]doc/1.jpg[/image]";
            pageofficeCtrl.SetWriter(doc);
            pageofficeCtrl.Caption = "演示：后台编程插入图片到数据区域并设置图片大小(企业版)";
            //隐藏菜单栏
            pageofficeCtrl.Menubar = false;
            pageofficeCtrl.CustomToolbar = false;

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}