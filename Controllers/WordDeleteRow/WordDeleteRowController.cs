using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordDeleteRow
{
    public class WordDeleteRowController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.Table table1 = doc.OpenDataRegion("PO_table").OpenTable(1);
            PageOfficeNetCore.WordWriter.Cell cell = table1.OpenCellRC(2, 1);
            //删除坐标为(2,1)的单元格所在行
            table1.RemoveRowAt(cell);
            pageofficeCtrl.CustomToolbar = false;
            pageofficeCtrl.SetWriter(doc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test_table.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}