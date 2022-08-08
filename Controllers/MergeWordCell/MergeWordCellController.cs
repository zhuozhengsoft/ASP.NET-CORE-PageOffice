using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.MergeWordCell
{
    public class MergeWordCellController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion dataReg = doc.OpenDataRegion("PO_table");
            PageOfficeNetCore.WordWriter.Table table = dataReg.OpenTable(1);
            //合并table中的单元格
            table.OpenCellRC(1, 1).MergeTo(1, 4);
            //给合并后的单元格赋值
            table.OpenCellRC(1, 1).Value = "销售情况表";
            //设置单元格文本样式
            table.OpenCellRC(1, 1).Font.Color = Color.Red;
            table.OpenCellRC(1, 1).Font.Size = 24;
            table.OpenCellRC(1, 1).Font.Name = "楷体";
            table.OpenCellRC(1, 1).ParagraphFormat.Alignment = PageOfficeNetCore.WordWriter.WdParagraphAlignment.wdAlignParagraphCenter;

            pageofficeCtrl.SetWriter(doc);//不要忘记此句代码
            pageofficeCtrl.CustomToolbar = false;

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}