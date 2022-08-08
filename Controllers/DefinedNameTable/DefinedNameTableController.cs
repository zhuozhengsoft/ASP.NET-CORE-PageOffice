using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DefinedNameTable
{
    public class DefinedNameTableController : Controller
    {

        private readonly IWebHostEnvironment _webHostEnvironment;

        public DefinedNameTableController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult ExcelFill()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook wk = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wk.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTableByDefinedName("report", 10, 5, false);
            table.DataFields[0].Value = "轮胎";
            table.DataFields[1].Value = "100";
            table.DataFields[2].Value = "120";
            table.DataFields[3].Value = "500";
            table.DataFields[4].Value = "120%";
            table.NextRow();
            table.Close();
            pageofficeCtrl.SetWriter(wk);// 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。

            pageofficeCtrl.Caption = "给Excel文档中定义名称的单元格赋值";
            pageofficeCtrl.SaveDataPage = "SaveData";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
        public IActionResult ExcelFill2()
        {

            string tempFileName = Request.Query["temp"];

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook wk = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wk.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTableByDefinedName("report", 10, 5, false);
            table.DataFields[0].Value = "轮胎";
            table.DataFields[1].Value = "100";
            table.DataFields[2].Value = "120";
            table.DataFields[3].Value = "500";
            table.DataFields[4].Value = "120%";
            table.NextRow();
            table.Close();
            // 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。
            //定义单元格对象，参数“year”就是Excel模板中定义的单元格的名称
            PageOfficeNetCore.ExcelWriter.Cell cellYear = sheet.OpenCellByDefinedName("year");
            // 给单元格赋值
            cellYear.Value = "2015年";

            PageOfficeNetCore.ExcelWriter.Cell cellName = sheet.OpenCellByDefinedName("name");
            cellName.Value = "张三";

            pageofficeCtrl.SetWriter(wk);
            //隐藏菜单栏
            pageofficeCtrl.Menubar = false;

            pageofficeCtrl.Caption = "给Excel文档中定义名称的单元格赋值";
            pageofficeCtrl.SaveDataPage = "SaveData";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/" + tempFileName, PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public IActionResult ExcelFill4()
        {

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            PageOfficeNetCore.ExcelWriter.Workbook wk = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wk.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("B4:F11");
            int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
            for (int i = 1; i <= rowCount; i++)
            {
                table.DataFields[0].Value = i + "月";
                table.DataFields[1].Value = "100";
                table.DataFields[2].Value = "120";
                table.DataFields[3].Value = "500";
                table.DataFields[4].Value = "120%";
                table.NextRow();
            }
            table.Close();

            //定义另一个table
            PageOfficeNetCore.ExcelWriter.Table table2 = sheet.OpenTable("B13:F16");
            int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
            for (int i = 1; i <= rowCount2; i++)
            {
                table2.DataFields[0].Value = i + "季度";
                table2.DataFields[1].Value = "300";
                table2.DataFields[2].Value = "300";
                table2.DataFields[3].Value = "300";
                table2.DataFields[4].Value = "100%";
                table2.NextRow();
            }

            table2.Close();
            pageofficeCtrl.SetWriter(wk);// 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。

            //隐藏菜单栏
            pageofficeCtrl.Menubar = false;
            pageofficeCtrl.Caption = "给Excel文档中定义名称的单元格赋值";
            pageofficeCtrl.SaveDataPage = "SaveData";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test4.xls", PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult ExcelFill5()
        {

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            PageOfficeNetCore.ExcelWriter.Workbook wk = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wk.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTableByDefinedName("report", 4, 5, true);
            int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
            for (int i = 1; i <= rowCount; i++)
            {
                table.DataFields[0].Value = i + "月";
                table.DataFields[1].Value = "100";
                table.DataFields[2].Value = "120";
                table.DataFields[3].Value = "500";
                table.DataFields[4].Value = "120%";
                table.NextRow();
            }

            table.Close();

            //定义另一个table
            PageOfficeNetCore.ExcelWriter.Table table2 = sheet.OpenTableByDefinedName("report2", 4, 5, true);
            int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
            for (int i = 1; i <= rowCount2; i++)
            {
                table2.DataFields[0].Value = i + "季度";
                table2.DataFields[1].Value = "300";
                table2.DataFields[2].Value = "300";
                table2.DataFields[3].Value = "300";
                table2.DataFields[4].Value = "100%";
                table2.NextRow();
            }

            table2.Close();
            pageofficeCtrl.SetWriter(wk);// 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。
            pageofficeCtrl.Caption = "给Excel文档中定义名称的单元格赋值";
            pageofficeCtrl.SaveDataPage = "SaveData";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test4.xls", PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
        public IActionResult ExcelFill6()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            pageofficeCtrl.Caption = "简单的给Excel赋值";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test4.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveData()
        {
            String content = "";
            PageOfficeNetCore.ExcelReader.Workbook doc = new PageOfficeNetCore.ExcelReader.Workbook(Request, Response);

            await doc.LoadAsync();
            PageOfficeNetCore.ExcelReader.Sheet sheet = doc.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelReader.Table table = sheet.OpenTable("report");
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
            // await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));
            doc.ShowPage(800, 800,this);
            doc.Close();
            ViewBag.content = content;
            return View();
        }
    }
}