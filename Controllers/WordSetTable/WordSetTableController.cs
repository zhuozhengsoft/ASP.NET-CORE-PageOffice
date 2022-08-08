using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordSetTable
{
    public class WordSetTableController : Controller
    {

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //获取Table所在的数据区域对象
            PageOfficeNetCore.WordWriter.DataRegion dataRegion = doc.OpenDataRegion("PO_regTable");
            //打开table，OpenTable(index)方法中的index代表Word文档中table位置的索引，从1开始
            PageOfficeNetCore.WordWriter.Table table = dataRegion.OpenTable(1);
            //给table中的单元格赋值， OpenCellRC(行, 列)
            table.OpenCellRC(3, 1).Value = "A公司";
            table.OpenCellRC(3, 2).Value = "开发部";
            table.OpenCellRC(3, 3).Value = "李清";
            //插入一空行，InsertRowAfter方法中的参数表示在哪个单元格下面插入一行
            table.InsertRowAfter(table.OpenCellRC(3, 3));

            table.OpenCellRC(4, 1).Value = "B公司";
            table.OpenCellRC(4, 2).Value = "销售部";
            table.OpenCellRC(4, 3).Value = "张三";
            pageofficeCtrl.SetWriter(doc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


    }
}