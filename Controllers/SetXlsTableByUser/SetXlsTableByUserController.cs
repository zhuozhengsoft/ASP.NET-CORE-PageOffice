using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SetXlsTableByUser
{
    public class SetXlsTableByUserController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        private readonly IWebHostEnvironment _webHostEnvironment;

        public SetXlsTableByUserController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Excel()
        {

            string userName = Request.Form["userName"];

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //***************************卓正PageOffice组件的使用********************************
            PageOfficeNetCore.ExcelWriter.Workbook wb = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = wb.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelWriter.Table tableA = sheet.OpenTable("C4:D6");
            PageOfficeNetCore.ExcelWriter.Table tableB = sheet.OpenTable("C7:D9");

            tableA.SubmitName = "tableA";
            tableB.SubmitName = "tableB";

            //根据登录用户名设置数据区域可编辑性
            //A部门经理登录后
            String strInfo = "";
            if (userName.Equals("zhangsan"))
            {
                strInfo = "A部门经理，所以只能编辑A部门的产品数据";
                tableA.ReadOnly = false;
                tableB.ReadOnly = true;
            }
            //B部门经理登录后
            else
            {
                strInfo = "B部门经理，所以只能编辑B部门的产品数据";
                tableA.ReadOnly = true;
                tableB.ReadOnly = false;
            }

            pageofficeCtrl.SetWriter(wb);

            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";

            pageofficeCtrl.SaveDataPage = "SaveData";//保存数据

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.strInfo = strInfo;
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SetXlsTableByUser/doc/" + fs.FileName);
            return fs.Close();
            
        }

        public async Task<ActionResult> SaveData()
        {

            PageOfficeNetCore.ExcelReader.Workbook doc = new PageOfficeNetCore.ExcelReader.Workbook(Request, Response);
            await doc.LoadAsync();
            PageOfficeNetCore.ExcelReader.Sheet sheet = doc.OpenSheet("Sheet1");
            PageOfficeNetCore.ExcelReader.Table tableA = sheet.OpenTable("tableA");
            PageOfficeNetCore.ExcelReader.Table tableB = sheet.OpenTable("tableB");

            StringBuilder dataStr = new StringBuilder();
            dataStr.Append("提交的数据为：<br/><br/>");
            dataStr.Append("<div style='float:left;width:460px;'>");
            dataStr.Append("<div style='float:left;width:150px;'>&nbsp; </div>");
            dataStr.Append("<div style='float:left;width:150px;'>计划完成量</div>");
            dataStr.Append("<div style='float:left;width:150px;'>实际完成量 </div>");
            dataStr.Append("</div>");
            while (!tableA.EOF)
            {
                dataStr.Append("<div style='float:left;width:460px;'>");
                dataStr.Append("<div style='float:left;width:150px;'> A部门：</div>");
                for (int i = 0; i < tableA.DataFields.Count; i++)
                {
                    dataStr.Append("<div style='float:left;width:150px;'>" + tableA.DataFields[i].Value + "</div>");
                }
                dataStr.Append("</div>");
                tableA.NextRow();
            }
            while (!tableB.EOF)
            {
                dataStr.Append("<div style='float:left;width:460px;'>");
                dataStr.Append("<div style='float:left;width:150px;'> B部门：</div>");
                for (int i = 0; i < tableB.DataFields.Count; i++)
                {
                    dataStr.Append("<div style='float:left;width:150px;'>" + tableB.DataFields[i].Value + "</div>");
                }
                dataStr.Append("</div>");
                tableB.NextRow();
            }
            // await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(dataStr.ToString()));
            //向客户端显示提交的数据
            doc.ShowPage(500, 400,this);
            doc.Close();
            ViewBag.dataStr = dataStr.ToString();
            return View();

        }

    }
}