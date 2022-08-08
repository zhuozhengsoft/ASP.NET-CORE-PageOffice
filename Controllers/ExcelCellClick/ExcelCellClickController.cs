using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExcelCellClick
{
    public class ExcelCellClickController : Controller
    {

        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");
            //定义table对象，设置table对象的设置范围
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("B4:D8");
            //设置table对象的提交名称，以便保存页面获取提交的数据
            table.SubmitName = "Info";
            pageofficeCtrl.SetWriter(workBook);

            // 设置响应单元格点击事件的js function
            pageofficeCtrl.JsFunction_OnExcelCellClick = "OnCellClick()";

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveData";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult select()
        {
            return View();
        }

        public async Task<ActionResult> SaveData()
        {
            PageOfficeNetCore.ExcelReader.Workbook doc = new PageOfficeNetCore.ExcelReader.Workbook(Request, Response);
            await doc.LoadAsync();
            PageOfficeNetCore.ExcelReader.Sheet sheet = doc.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelReader.Table table = sheet.OpenTable("B4:D8");
            String content = "";
            while (!table.EOF)
            {
                //获取提交的数值
                //DataFields.Count标识的是table的列数
                if (!table.DataFields.IsEmpty)
                {
                    content += "<br/>月份名称：" + table.DataFields[0].Text;
                    content += "<br/>计划完成量：" + table.DataFields[1].Text;
                    content += "<br/>实际完成量：" + table.DataFields[2].Text;

                    content += "<br/>*********************************************";
                }
                //循环进入下一行
                table.NextRow();
            }
            table.Close();
            //  await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));
            doc.ShowPage(500, 400,this);
            doc.Close();
            ViewBag.content = content;
            return View();
        }
    }
}