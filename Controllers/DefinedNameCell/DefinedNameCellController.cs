using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DefinedNameCell
{
    public class DefinedNameCellController : Controller
    {
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //定义Workbook对象
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");
            sheet.OpenCellByDefinedName("testA1").Value = "Tom";
            sheet.OpenCellByDefinedName("testB1").Value = "John";
            // 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。
            pageofficeCtrl.SetWriter(workBook);

            pageofficeCtrl.Caption = "给Excel文档中定义名称的单元格赋值";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveData";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveData()
        {
            string content = "";
                   
            PageOfficeNetCore.ExcelReader.Workbook workBook = new PageOfficeNetCore.ExcelReader.Workbook(Request, Response);
            await workBook.LoadAsync();

            PageOfficeNetCore.ExcelReader.Sheet sheet = workBook.OpenSheet("Sheet1");
            content += "testA1：" + sheet.OpenCellByDefinedName("testA1").Value + "<br/>";
            content += "testB1：" + sheet.OpenCellByDefinedName("testB1").Value + "<br/>";
            //  await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));
            workBook.ShowPage(500, 400,this);
            workBook.Close();

            ViewBag.content = content;
            return View();
        }
    }
}