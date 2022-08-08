using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ImportWordData
{
    public class ImportWordDataController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public ImportWordDataController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.AddCustomToolButton("导入文件", "importData()", 5);
            pageofficeCtrl.AddCustomToolButton("提交数据", "submitData()", 1);
            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();
            pageofficeCtrl.SetWriter(wordDoc);

            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveDoc";
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            String conment = "";

            string docID = Request.Query["id"];

            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);

            await doc.LoadAsync();

            String sName = doc.OpenDataRegion("PO_name").Value;
            String sDept = doc.OpenDataRegion("PO_dept").Value;
            String sCause = doc.OpenDataRegion("PO_cause").Value;
            String sNum = doc.OpenDataRegion("PO_num").Value;
            String sDate = doc.OpenDataRegion("PO_date").Value;

            conment += "提交的数据为：<br/>";
            conment += "姓名：" + sName + "<br/>";
            conment += "原因：" + sCause + "<br/>";
            conment += "天数：" + sNum + "<br/>";

            conment += "日期：" + sDate + "<br/>";

            //await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(conment));

            doc.ShowPage(578, 380,this);
            doc.Close();
            ViewBag.conment = conment;
            return View();
        }
    }
}