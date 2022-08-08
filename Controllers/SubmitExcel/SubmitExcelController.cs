using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SubmitExcel
{
    public class SubmitExcelController : Controller
    {

        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            //定义Workbook对象
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");

            //定义table对象，设置table对象的设置范围
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("B4:F13");
            //设置table对象的提交名称，以便保存页面获取提交的数据
            table.SubmitName = "Info";

            pageofficeCtrl.SetWriter(workBook);
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
            PageOfficeNetCore.ExcelReader.Table table = sheet.OpenTable("Info");
            int result = 0;
            while (!table.EOF)
            {
                //获取提交的数值
                //DataFields.Count标识的是提交过来的table的列数
                if (!table.DataFields.IsEmpty)
                {
                    content += "<br/>月份名称：" + table.DataFields[0].Text;
                    content += "<br/>计划完成量：" + table.DataFields[1].Text;
                    content += "<br/>实际完成量：" + table.DataFields[2].Text;
                    content += "<br/>累计完成量：" + table.DataFields[3].Text;
                    if (string.IsNullOrEmpty(table.DataFields[2].Text) || !int.TryParse(table.DataFields[2].Text, out result) ||
                        !int.TryParse(table.DataFields[1].Text, out result))
                    {
                        content += "<br/>完成率：0";
                    }
                    else
                    {
                        float f = int.Parse(table.DataFields[2].Text);
                        f = f / int.Parse(table.DataFields[1].Text);
                        content += "<br/>完成率：" + string.Format("{0:P}", f);
                    }
                    content += "<br/>*********************************************";
                }
                //循环进入下一行
                table.NextRow();
            }
            table.Close();
            //await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));
            workBook.ShowPage(500, 400,this);
            workBook.Close();
            ViewBag.content = content;
            return View();
        }
    }
}