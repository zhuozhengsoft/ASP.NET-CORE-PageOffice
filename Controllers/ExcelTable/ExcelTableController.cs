using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExcelTable
{
    public class ExcelTableController : Controller
    {
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //定义Workbook对象
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");
            //定义Table对象
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("B4:F13");
            for (int i = 0; i < 50; i++)
            {
                table.DataFields[0].Value = "产品 " + i.ToString();
                table.DataFields[1].Value = "100";
                table.DataFields[2].Value = (100 + i).ToString();
                table.NextRow();
            }
            table.Close();

            pageofficeCtrl.SetWriter(workBook);// 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。
            pageofficeCtrl.Caption = "使用OpenTable给Excel赋值";

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}