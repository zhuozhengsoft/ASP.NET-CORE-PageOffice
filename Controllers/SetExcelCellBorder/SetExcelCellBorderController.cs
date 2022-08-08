using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SetExcelCellBorder
{
    public class SetExcelCellBorderController : Controller
    {
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook wb = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wb.OpenSheet("Sheet1");
            // 设置背景
            PageOfficeNetCore.ExcelWriter.Table backGroundTable = sheet.OpenTable("A1:P200");
            //设置表格边框样式
            backGroundTable.Border.LineColor = Color.White;

            // 设置单元格边框样式
            PageOfficeNetCore.ExcelWriter.Border C4Border = sheet.OpenTable("C4:C4").Border;
            C4Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            C4Border.LineColor = Color.Yellow;
            C4Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;

            // 设置单元格边框样式
            PageOfficeNetCore.ExcelWriter.Border B6Border = sheet.OpenTable("B6:B6").Border;
            B6Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlHairline;
            B6Border.LineColor = Color.Purple;
            B6Border.LineStyle = PageOfficeNetCore.ExcelWriter.XlBorderLineStyle.xlSlantDashDot;
            B6Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;

            //设置表格边框样式
            PageOfficeNetCore.ExcelWriter.Table titleTable = sheet.OpenTable("B4:F5");
            titleTable.Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            titleTable.Border.LineColor = Color.FromArgb(0, 128, 128);
            titleTable.Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;

            //设置表格边框样式
            PageOfficeNetCore.ExcelWriter.Table bodyTable2 = sheet.OpenTable("B6:F15");
            bodyTable2.Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            bodyTable2.Border.LineColor = Color.FromArgb(0, 128, 128);
            bodyTable2.Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;

            pageofficeCtrl.SetWriter(wb);// 不要忘记此句代码

            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}