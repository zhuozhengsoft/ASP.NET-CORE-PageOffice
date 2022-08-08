using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DrawExcel
{
    public class DrawExcelController : Controller
    {

        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook wb = new PageOfficeNetCore.ExcelWriter.Workbook();
            // 设置背景
            PageOfficeNetCore.ExcelWriter.Table backGroundTable = wb.OpenSheet("Sheet1").OpenTable("A1:P200");
            backGroundTable.Border.LineColor = Color.White;

            // 设置标题
            wb.OpenSheet("Sheet1").OpenTable("A1:H2").Merge();
            wb.OpenSheet("Sheet1").OpenTable("A1:H2").RowHeight = 30;
            PageOfficeNetCore.ExcelWriter.Cell A1 = wb.OpenSheet("Sheet1").OpenCell("A1");
            A1.HorizontalAlignment = PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            A1.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;
            A1.ForeColor = Color.FromArgb(0, 128, 128);
            A1.Value = "出差开支预算";
            A1.Font.Bold = true;
            A1.Font.Size = 25;

            #region 画表头
            // 画表头
            PageOfficeNetCore.ExcelWriter.Border C4Border = wb.OpenSheet("Sheet1").OpenTable("C4:C4").Border;
            C4Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            C4Border.LineColor = Color.Yellow;

            PageOfficeNetCore.ExcelWriter.Table titleTable = wb.OpenSheet("Sheet1").OpenTable("B4:H5");
            titleTable.Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            titleTable.Border.LineColor = Color.FromArgb(0, 128, 128);
            titleTable.Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;
            #endregion

            #region 画表体
            // 画表体
            PageOfficeNetCore.ExcelWriter.Table bodyTable = wb.OpenSheet("Sheet1").OpenTable("B6:H15");
            bodyTable.Border.LineColor = Color.Gray;
            bodyTable.Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlHairline;

            PageOfficeNetCore.ExcelWriter.Border B7Border = wb.OpenSheet("Sheet1").OpenTable("B7:B7").Border;
            B7Border.LineColor = Color.White;

            PageOfficeNetCore.ExcelWriter.Border B9Border = wb.OpenSheet("Sheet1").OpenTable("B9:B9").Border;
            B9Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlBottomEdge;
            B9Border.LineColor = Color.White;

            PageOfficeNetCore.ExcelWriter.Border C6C15BorderLeft = wb.OpenSheet("Sheet1").OpenTable("C6:C15").Border;
            C6C15BorderLeft.LineColor = Color.White;
            C6C15BorderLeft.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlLeftEdge;
            PageOfficeNetCore.ExcelWriter.Border C6C15BorderRight = wb.OpenSheet("Sheet1").OpenTable("C6:C15").Border;
            C6C15BorderRight.LineColor = Color.Yellow;
            C6C15BorderRight.LineStyle = PageOfficeNetCore.ExcelWriter.XlBorderLineStyle.xlDot;
            C6C15BorderRight.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlRightEdge;

            PageOfficeNetCore.ExcelWriter.Border E6E15Border = wb.OpenSheet("Sheet1").OpenTable("E6:E15").Border;
            E6E15Border.LineStyle = PageOfficeNetCore.ExcelWriter.XlBorderLineStyle.xlDot;
            E6E15Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;
            E6E15Border.LineColor = Color.Yellow;

            PageOfficeNetCore.ExcelWriter.Border G6G15BorderRight = wb.OpenSheet("Sheet1").OpenTable("G6:G15").Border;
            G6G15BorderRight.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlRightEdge;
            G6G15BorderRight.LineColor = Color.White;
            PageOfficeNetCore.ExcelWriter.Border G6G15BorderLeft = wb.OpenSheet("Sheet1").OpenTable("G6:G15").Border;
            G6G15BorderLeft.LineStyle = PageOfficeNetCore.ExcelWriter.XlBorderLineStyle.xlDot;
            G6G15BorderLeft.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlLeftEdge;
            G6G15BorderLeft.LineColor = Color.Yellow;

            PageOfficeNetCore.ExcelWriter.Table bodyTable2 = wb.OpenSheet("Sheet1").OpenTable("B6:H15");
            bodyTable2.Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            bodyTable2.Border.LineColor = Color.FromArgb(0, 128, 128);
            bodyTable2.Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;
            #endregion

            #region 画表尾
            // 画表尾
            PageOfficeNetCore.ExcelWriter.Border H16H17Border = wb.OpenSheet("Sheet1").OpenTable("H16:H17").Border;
            H16H17Border.LineColor = Color.FromArgb(204, 255, 204);

            PageOfficeNetCore.ExcelWriter.Border E16G17Border = wb.OpenSheet("Sheet1").OpenTable("E16:G17").Border;
            E16G17Border.LineColor = Color.FromArgb(0, 128, 128);

            PageOfficeNetCore.ExcelWriter.Table footTable = wb.OpenSheet("Sheet1").OpenTable("B16:H17");
            footTable.Border.Weight = PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThick;
            footTable.Border.LineColor = Color.FromArgb(0, 128, 128);
            footTable.Border.BorderType = PageOfficeNetCore.ExcelWriter.XlBorderType.xlAllEdges;
            #endregion

            #region 设置行高列宽
            // 设置行高列宽
            wb.OpenSheet("Sheet1").OpenTable("A1:A1").ColumnWidth = 1;
            wb.OpenSheet("Sheet1").OpenTable("B1:B1").ColumnWidth = 20;
            wb.OpenSheet("Sheet1").OpenTable("C1:C1").ColumnWidth = 15;
            wb.OpenSheet("Sheet1").OpenTable("D1:D1").ColumnWidth = 10;
            wb.OpenSheet("Sheet1").OpenTable("E1:E1").ColumnWidth = 8;
            wb.OpenSheet("Sheet1").OpenTable("F1:F1").ColumnWidth = 3;
            wb.OpenSheet("Sheet1").OpenTable("G1:G1").ColumnWidth = 12;
            wb.OpenSheet("Sheet1").OpenTable("H1:H1").ColumnWidth = 20;

            wb.OpenSheet("Sheet1").OpenTable("A16:A16").RowHeight = 20;
            wb.OpenSheet("Sheet1").OpenTable("A17:A17").RowHeight = 20;
            #endregion

            // 设置表格中字体大小为10
            for (int i = 0; i < 12; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    wb.OpenSheet("Sheet1").OpenCellRC(4 + i, 2 + j).Font.Size = 10;
                }
            }

            #region 填充单元格背景颜色

            // 填充单元格背景颜色
            for (int i = 0; i < 10; i++)
            {
                wb.OpenSheet("Sheet1").OpenCell("H" + (6 + i).ToString()).BackColor = Color.FromArgb(255, 255, 153);
            }

            wb.OpenSheet("Sheet1").OpenCell("E16").BackColor = Color.FromArgb(0, 128, 128);
            wb.OpenSheet("Sheet1").OpenCell("F16").BackColor = Color.FromArgb(0, 128, 128);
            wb.OpenSheet("Sheet1").OpenCell("G16").BackColor = Color.FromArgb(0, 128, 128);
            wb.OpenSheet("Sheet1").OpenCell("E17").BackColor = Color.FromArgb(0, 128, 128);
            wb.OpenSheet("Sheet1").OpenCell("F17").BackColor = Color.FromArgb(0, 128, 128);
            wb.OpenSheet("Sheet1").OpenCell("G17").BackColor = Color.FromArgb(0, 128, 128);

            wb.OpenSheet("Sheet1").OpenCell("H16").BackColor = Color.FromArgb(204, 255, 204);
            wb.OpenSheet("Sheet1").OpenCell("H17").BackColor = Color.FromArgb(204, 255, 204);
            #endregion

            #region 填充单元格文本和公式

            //填充单元格文本和公式
            PageOfficeNetCore.ExcelWriter.Cell B4 = wb.OpenSheet("Sheet1").OpenCell("B4");
            B4.Font.Bold = true;
            B4.Value = "出差开支预算";
            PageOfficeNetCore.ExcelWriter.Cell H5 = wb.OpenSheet("Sheet1").OpenCell("H5");
            H5.Font.Bold = true;
            H5.Value = "总计";
            H5.HorizontalAlignment = PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            PageOfficeNetCore.ExcelWriter.Cell B6 = wb.OpenSheet("Sheet1").OpenCell("B6");
            B6.Font.Bold = true;
            B6.Value = "飞机票价";
            PageOfficeNetCore.ExcelWriter.Cell B9 = wb.OpenSheet("Sheet1").OpenCell("B9");
            B9.Font.Bold = true;
            B9.Value = "酒店";
            PageOfficeNetCore.ExcelWriter.Cell B11 = wb.OpenSheet("Sheet1").OpenCell("B11");
            B11.Font.Bold = true;
            B11.Value = "餐饮";
            PageOfficeNetCore.ExcelWriter.Cell B12 = wb.OpenSheet("Sheet1").OpenCell("B12");
            B12.Font.Bold = true;
            B12.Value = "交通费用";
            PageOfficeNetCore.ExcelWriter.Cell B13 = wb.OpenSheet("Sheet1").OpenCell("B13");
            B13.Font.Bold = true;
            B13.Value = "休闲娱乐";
            PageOfficeNetCore.ExcelWriter.Cell B14 = wb.OpenSheet("Sheet1").OpenCell("B14");
            B14.Font.Bold = true;
            B14.Value = "礼品";
            PageOfficeNetCore.ExcelWriter.Cell B15 = wb.OpenSheet("Sheet1").OpenCell("B15");
            B15.Font.Bold = true;
            B15.Font.Size = 10;
            B15.Value = "其他费用";

            wb.OpenSheet("Sheet1").OpenCell("C6").Value = "机票单价（往）";
            wb.OpenSheet("Sheet1").OpenCell("C7").Value = "机票单价（返）";
            wb.OpenSheet("Sheet1").OpenCell("C8").Value = "其他";
            wb.OpenSheet("Sheet1").OpenCell("C9").Value = "每晚费用";
            wb.OpenSheet("Sheet1").OpenCell("C10").Value = "其他";
            wb.OpenSheet("Sheet1").OpenCell("C11").Value = "每天费用";
            wb.OpenSheet("Sheet1").OpenCell("C12").Value = "每天费用";
            wb.OpenSheet("Sheet1").OpenCell("C13").Value = "总计";
            wb.OpenSheet("Sheet1").OpenCell("C14").Value = "总计";
            wb.OpenSheet("Sheet1").OpenCell("C15").Value = "总计";

            wb.OpenSheet("Sheet1").OpenCell("G6").Value = "  张";
            wb.OpenSheet("Sheet1").OpenCell("G7").Value = "  张";
            wb.OpenSheet("Sheet1").OpenCell("G9").Value = "  晚";
            wb.OpenSheet("Sheet1").OpenCell("G10").Value = "  晚";
            wb.OpenSheet("Sheet1").OpenCell("G11").Value = "  天";
            wb.OpenSheet("Sheet1").OpenCell("G12").Value = "  天";

            wb.OpenSheet("Sheet1").OpenCell("H6").Formula = "=D6*F6";
            wb.OpenSheet("Sheet1").OpenCell("H7").Formula = "=D7*F7";
            wb.OpenSheet("Sheet1").OpenCell("H8").Formula = "=D8*F8";
            wb.OpenSheet("Sheet1").OpenCell("H9").Formula = "=D9*F9";
            wb.OpenSheet("Sheet1").OpenCell("H10").Formula = "=D10*F10";
            wb.OpenSheet("Sheet1").OpenCell("H11").Formula = "=D11*F11";
            wb.OpenSheet("Sheet1").OpenCell("H12").Formula = "=D12*F12";
            wb.OpenSheet("Sheet1").OpenCell("H13").Formula = "=D13*F13";
            wb.OpenSheet("Sheet1").OpenCell("H14").Formula = "=D14*F14";
            wb.OpenSheet("Sheet1").OpenCell("H15").Formula = "=D15*F15";

            for (int i = 0; i < 10; i++)
            {
                wb.OpenSheet("Sheet1").OpenCell("D" + (6 + i).ToString()).NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
                wb.OpenSheet("Sheet1").OpenCell("H" + (6 + i).ToString()).NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
            }

            PageOfficeNetCore.ExcelWriter.Cell E16 = wb.OpenSheet("Sheet1").OpenCell("E16");
            E16.Font.Bold = true;
            E16.Font.Size = 11;
            E16.ForeColor = Color.White;
            E16.Value = "出差开支总费用";
            E16.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;
            PageOfficeNetCore.ExcelWriter.Cell E17 = wb.OpenSheet("Sheet1").OpenCell("E17");
            E17.Font.Bold = true;
            E17.Font.Size = 11;
            E17.ForeColor = Color.White;
            E17.Formula = "=IF(C4>H16,\"低于预算\",\"超出预算\")";
            E17.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;
            PageOfficeNetCore.ExcelWriter.Cell H16 = wb.OpenSheet("Sheet1").OpenCell("H16");
            H16.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;
            H16.NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
            H16.Font.Name = "Arial";
            H16.Font.Size = 11;
            H16.Font.Bold = true;
            H16.Formula = "=SUM(H6:H15)";
            PageOfficeNetCore.ExcelWriter.Cell H17 = wb.OpenSheet("Sheet1").OpenCell("H17");
            H17.VerticalAlignment = PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;
            H17.NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
            H17.Font.Name = "Arial";
            H17.Font.Size = 11;
            H17.Font.Bold = true;
            H17.Formula = "=(C4-H16)";
            #endregion

            #region 填充数据
            // 填充数据
            PageOfficeNetCore.ExcelWriter.Cell C4 = wb.OpenSheet("Sheet1").OpenCell("C4");
            C4.NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
            C4.Value = "2500";
            PageOfficeNetCore.ExcelWriter.Cell D6 = wb.OpenSheet("Sheet1").OpenCell("D6");
            D6.NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
            D6.Value = "1200";
            wb.OpenSheet("Sheet1").OpenCell("F6").Font.Size = 10;
            wb.OpenSheet("Sheet1").OpenCell("F6").Value = "1";
            PageOfficeNetCore.ExcelWriter.Cell D7 = wb.OpenSheet("Sheet1").OpenCell("D7");
            D7.NumberFormatLocal = "￥#,##0.00;￥-#,##0.00";
            D7.Value = "875";
            wb.OpenSheet("Sheet1").OpenCell("F7").Value = "1";

            #endregion

            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            pageofficeCtrl.Caption = "完全使用程序生成Excel表格";
            // 打开文件

            pageofficeCtrl.SetWriter(wb);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}