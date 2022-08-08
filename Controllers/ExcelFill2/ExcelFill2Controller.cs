using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExcelFill2
{
    public class ExcelFill2Controller : Controller
    {
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            pageofficeCtrl.Caption = "简单的给Excel赋值";
            //定义Workbook对象
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");
            //定义Cell对象
            PageOfficeNetCore.ExcelWriter.Cell cellB4 = sheet.OpenCell("B4");
            //给单元格赋值
            cellB4.Value = "1月";
            //设置字体颜色
            cellB4.ForeColor = Color.Red;

            PageOfficeNetCore.ExcelWriter.Cell cellC4 = sheet.OpenCell("C4");
            cellC4.Value = "300";
            cellC4.ForeColor = Color.Blue;

            PageOfficeNetCore.ExcelWriter.Cell cellD4 = sheet.OpenCell("D4");
            cellD4.Value = "270";
            cellD4.ForeColor = Color.Orange;

            PageOfficeNetCore.ExcelWriter.Cell cellE4 = sheet.OpenCell("E4");
            cellE4.Value = "270";
            cellE4.ForeColor = Color.Green;

            PageOfficeNetCore.ExcelWriter.Cell cellF4 = sheet.OpenCell("F4");
            cellF4.Value = string.Format("{0:P}", 270.0 / 300);
            cellF4.ForeColor = Color.Gray;

            pageofficeCtrl.SetWriter(workBook);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}