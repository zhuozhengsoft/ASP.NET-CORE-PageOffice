using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.MergeExcelCell
{
    public class MergeExcelCellController : Controller
    {
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook wb = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wb.OpenSheet("Sheet1");
            //合并单元格
            sheet.OpenTable("B2:F2").Merge();
            PageOfficeNetCore.ExcelWriter.Cell cB2 = sheet.OpenCell("B2");
            cB2.Value = "北京某公司产品销售情况";
            //设置水平对齐方式
            cB2.HorizontalAlignment = PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            cB2.ForeColor = Color.Red;
            cB2.Font.Size = 16;

            sheet.OpenTable("B4:B6").Merge();//合并单元格
            PageOfficeNetCore.ExcelWriter.Cell cB4 = sheet.OpenCell("B4");
            cB4.Value = "A产品";
            //设置水平对齐方式
            cB4.HorizontalAlignment = PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            //设置垂直对齐方式
            cB4.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;

            sheet.OpenTable("B7:B9").Merge();//合并单元格
            PageOfficeNetCore.ExcelWriter.Cell cB7 = sheet.OpenCell("B7");
            cB7.Value = "B产品";
            cB7.HorizontalAlignment = PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            cB7.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;

            pageofficeCtrl.SetWriter(wb);

            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            pageofficeCtrl.Caption = "演示：使用程序合并指定的单元格并设置格式和赋值";

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}