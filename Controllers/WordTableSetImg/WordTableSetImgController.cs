using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordTableSetImg
{
    public class WordTableSetImgController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.Table table1 = doc.OpenDataRegion("PO_T001").OpenTable(1);
            table1.OpenCellRC(1, 1).Value = "[image]doc/wang.gif[/image]";

            int oldRowCount = 3;//表格中原有的行数
            int dataRowCount = 5;//要填充数据的行数
                                 // 扩充表格
            for (int j = 0; j < dataRowCount - oldRowCount; j++)
            {
                table1.InsertRowAfter(table1.OpenCellRC(2, 5));  //在第2行的最后一个单元格下插入新行
            }

            // 填充数据
            int i = 1;
            while (i <= dataRowCount)
            {
                table1.OpenCellRC(i, 2).Value = "AA" + i.ToString();
                table1.OpenCellRC(i, 3).Value = "BB" + i.ToString();
                table1.OpenCellRC(i, 4).Value = "CC" + i.ToString();
                table1.OpenCellRC(i, 5).Value = "DD" + i.ToString();
                i++;
            }
            pageofficeCtrl.SetWriter(doc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test_table.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}