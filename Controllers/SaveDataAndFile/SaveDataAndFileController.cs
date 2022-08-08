using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SaveDataAndFile
{
    public class SaveDataAndFileController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public SaveDataAndFileController(IWebHostEnvironment webHostEnvironment)
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

            PageOfficeNetCore.WordWriter.DataRegion dataRegion2 = wordDoc.OpenDataRegion("PO_deptName");
            dataRegion2.Editing = true;

            pageofficeCtrl.SetWriter(wordDoc);

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            //设置保存数据的页面
            pageofficeCtrl.SaveDataPage = "SaveData";
            //设置保存文件的页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SaveDataAndFile/doc/" + fs.FileName);
            return fs.Close();
            
        }

        public async Task<ActionResult> SaveData()
        {
            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();

            //获取提交的数值
            String dataUserName = doc.OpenDataRegion("PO_userName").Value;
            String dataDeptName = doc.OpenDataRegion("PO_deptName").Value;
            String companyName = doc.GetFormField("txtCompany");
            /**获取到的公司名称,员工姓名,部门名称等内容可以保存到数据库,这里可以连接开发人员自己的数据库,例如连接sqlServer2008数据库。
             * string constr = "server=ACER-PC\\LI;database=db_test;uid=sa;pwd=123";
             * conn = new SqlConnection(constr);  //数据库连接  
             * conn.Open();
             * SqlCommand cmd = new SqlCommand("update user set userName='"+dataUserName+"',deptName='"+dataDeptName+"',companyName='"+companyName+"' where userId=1",conn);
             * cmd.ExecuteNonQuery();
             * conn.Close();
             * */

            return doc.Close();
            
        }

    }
}