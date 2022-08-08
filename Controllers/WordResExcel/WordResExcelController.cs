using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordResExcel
{
    public class WordResExcelController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument worddoc = new PageOfficeNetCore.WordWriter.WordDocument();
            //先在要插入word文件的位置手动插入书签,书签必须以“PO_”为前缀
            //给DataRegion赋值,值的形式为："[word]word文件路径[/word]"
            PageOfficeNetCore.WordWriter.DataRegion data1 = worddoc.OpenDataRegion("PO_p1");
            data1.Value = "[excel]doc/1.xls[/excel]";
            PageOfficeNetCore.WordWriter.DataRegion data2 = worddoc.OpenDataRegion("PO_p2");
            data2.Value = "[word]doc/2.doc[/word]";
            PageOfficeNetCore.WordWriter.DataRegion data3 = worddoc.OpenDataRegion("PO_p3");
            data3.Value = "[word]doc/3.doc[/word]";

            pageofficeCtrl.SetWriter(worddoc);
            pageofficeCtrl.Caption = "演示：后台编程插入excel文件到数据区域(企业版)";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}