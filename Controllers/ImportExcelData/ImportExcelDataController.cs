using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ImportExcelData
{
    public class ImportExcelDataController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public ImportExcelDataController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //定义Workbook对象
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("Sheet1");
            pageofficeCtrl.SetWriter(workBook);

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("导入文件", "importData()", 5);
            pageofficeCtrl.AddCustomToolButton("提交数据", "submitData()", 1);

            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveDoc";
            //打开Word文档
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            string content = "";
            PageOfficeNetCore.ExcelReader.Workbook doc = new PageOfficeNetCore.ExcelReader.Workbook(Request, Response);
            await doc.LoadAsync();

            PageOfficeNetCore.ExcelReader.Sheet sheet = doc.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelReader.Table table = sheet.OpenTable("B4:F13");
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
                    content += "<br/>";
                }
                //循环进入下一行
                table.NextRow();
            }
            table.Close();
            doc.ShowPage(500, 400,this);

            //await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));
            doc.Close();
            ViewBag.content = content;
            return View();
        }
    }
}