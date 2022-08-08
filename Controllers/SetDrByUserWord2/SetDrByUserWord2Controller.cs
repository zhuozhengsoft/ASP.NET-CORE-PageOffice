using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SetDrByUserWord2
{
    public class SetDrByUserWord2Controller : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public SetDrByUserWord2Controller(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Word()
        {

            string userName = Request.Form["userName"];

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();

            //打开数据区域
            PageOfficeNetCore.WordWriter.DataRegion d1 = doc.OpenDataRegion("PO_com1");
            //给数据区域赋值
            d1.Value = "[word]doc/content1.doc[/word]";
            //若要将数据区域内容存入文件中，则必须设置属性“SubmitAsFile”值为true
            d1.SubmitAsFile = true;

            PageOfficeNetCore.WordWriter.DataRegion d2 = doc.OpenDataRegion("PO_com2");
            d2.Value = "[word]doc/content2.doc[/word]";
            d2.SubmitAsFile = true;

            //根据登录用户名设置数据区域可编辑性
            //甲客户：zhangsan 登录后登录后
            if (userName.Equals("zhangsan"))
            {
                d1.Editing = true;
                d2.Editing = false;
            }
            //乙客户：lisi 登录后登录后
            else
            {
                d2.Editing = true;
                d1.Editing = false;
            }

            pageofficeCtrl.SetWriter(doc);

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);

            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveData?userName=" + userName;
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, userName);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.userName = userName;
            return View();
        }

        public async Task<ActionResult> SaveData()
        {
            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;

            if (Request.Query["userName"] != "" && Request.Query["userName"].Equals("zhangsan"))
            {
                saveBytesToFile(doc.OpenDataRegion("PO_com1").FileBytes, webRootPath + "/SetDrByUserWord2/doc/content1.doc");
            }
            else
            {
                saveBytesToFile(doc.OpenDataRegion("PO_com2").FileBytes, webRootPath + "/SetDrByUserWord2/doc/content2.doc");
            }

            return doc.Close();
            
        }
        private void saveBytesToFile(byte[] bytes, string filePath)
        {
            FileStream fs = new FileStream(filePath, System.IO.FileMode.OpenOrCreate);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
        }

    }
}