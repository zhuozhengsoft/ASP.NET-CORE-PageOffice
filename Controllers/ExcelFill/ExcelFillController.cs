using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExcelFill
{
    public class ExcelFillController : Controller
    {

        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //定义Workbook对象
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");
            //定义Cell对象,给单元格赋值
            PageOfficeNetCore.ExcelWriter.Cell cellB4 = sheet.OpenCell("B4");
            cellB4.Value = "1月";
            //或者直接给Cell赋值
            sheet.OpenCell("C4").Value = "300";
            sheet.OpenCell("D4").Value = "270";
            sheet.OpenCell("E4").Value = "270";
            sheet.OpenCell("F4").Value = string.Format("{0:P}", 270.0 / 300);

            pageofficeCtrl.SetWriter(workBook);// 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。

            pageofficeCtrl.Caption = "简单的给Excel赋值";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


    }
}