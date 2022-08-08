using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SplitWord
{
    public class SplitWordController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public SplitWordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            // 设置保存文件页面
            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.CustomToolbar = true;
            //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
            PageOfficeNetCore.WordWriter.DataRegion dataRegion1 = wordDoc.OpenDataRegion("PO_test1");
            dataRegion1.SubmitAsFile = true;
            PageOfficeNetCore.WordWriter.DataRegion dataRegion2 = wordDoc.OpenDataRegion("PO_test2");
            dataRegion2.SubmitAsFile = true;
            dataRegion2.Editing = true;
            PageOfficeNetCore.WordWriter.DataRegion dataRegion3 = wordDoc.OpenDataRegion("PO_test3");
            dataRegion3.SubmitAsFile = true;

            pageofficeCtrl.SetWriter(wordDoc);

            pageofficeCtrl.SaveDataPage = "SaveData";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveData()
        {
            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            Byte[] bWord;

            // 读取数据区域PO_test1中的内容，保存为一个新的word文档：new1.doc
            PageOfficeNetCore.WordReader.DataRegion dr1 = doc.OpenDataRegion("PO_test1");
            bWord = dr1.FileBytes;

            Stream s1 = new FileStream(webRootPath + "/SplitWord/doc/new1.doc", System.IO.FileMode.Create);

            s1.Write(bWord, 0, bWord.Length);
            s1.Close();

            // 读取数据区域PO_test2中的内容，保存为一个新的word文档：new2.doc
            PageOfficeNetCore.WordReader.DataRegion dr2 = doc.OpenDataRegion("PO_test2");
            bWord = dr2.FileBytes;
            Stream s2 = new FileStream(webRootPath + "/SplitWord/doc/new2.doc", System.IO.FileMode.Create);
            s2.Write(bWord, 0, bWord.Length);
            s2.Close();

            // 读取数据区域PO_test3中的内容，保存为一个新的word文档：new3.doc
            PageOfficeNetCore.WordReader.DataRegion dr3 = doc.OpenDataRegion("PO_test3");
            bWord = dr3.FileBytes;
            Stream s3 = new FileStream(webRootPath + "/SplitWord/doc/new3.doc", FileMode.Create);
            s3.Write(bWord, 0, bWord.Length);
            s3.Close();

            return doc.Close();
            
        }
    }
}