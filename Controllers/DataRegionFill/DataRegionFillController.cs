using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using PageOfficeNetCore.WordReader;

namespace NetCoreSamples5.Controllers
{
    public class DataRegionFillController : Controller
    {

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //打开数据区域
            PageOfficeNetCore.WordWriter.DataRegion dataRegion1 = doc.OpenDataRegion("PO_userName");
            dataRegion1.Value = "张三";

            PageOfficeNetCore.WordWriter.DataRegion dataRegion2 = doc.OpenDataRegion("PO_deptName");
            dataRegion2.Value = "销售部";

            pageofficeCtrl.SetWriter(doc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}