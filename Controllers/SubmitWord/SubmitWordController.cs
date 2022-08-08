using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SubmitWord
{
    public class SubmitWordController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public SubmitWordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();

            //打开数据区域，OpenDataRegion方法的参数代表Word文档中的书签名称
            PageOfficeNetCore.WordWriter.DataRegion dataRegion1 = wordDoc.OpenDataRegion("PO_userName");
            //设置DataRegion的可编辑性
            dataRegion1.Editing = true;
            //为DataRegion赋值,此处的值可在页面中打开Word文档后在自己进行修改
            dataRegion1.Value = "";

            PageOfficeNetCore.WordWriter.DataRegion dataRegion2 = wordDoc.OpenDataRegion("PO_deptName");
            dataRegion2.Editing = true;
            dataRegion2.Value = "";

            pageofficeCtrl.SetWriter(wordDoc);
            pageofficeCtrl.SaveDataPage = "SaveData";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveData()
        {

            string content = "";


            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();

            //获取提交的数值
            PageOfficeNetCore.WordReader.DataRegion dataUserName = doc.OpenDataRegion("PO_userName");
            PageOfficeNetCore.WordReader.DataRegion dataDeptName = doc.OpenDataRegion("PO_deptName");
            content += "公司名称：" + doc.GetFormField("txtCompany");
            content += "<br/>员工姓名：" + dataUserName.Value;
            content += "<br/>部门名称：" + dataDeptName.Value;

           // await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(content));

            doc.ShowPage(400, 300,this);
            doc.Close();
            ViewBag.content = content;
            return View() ;
        }
    }
}