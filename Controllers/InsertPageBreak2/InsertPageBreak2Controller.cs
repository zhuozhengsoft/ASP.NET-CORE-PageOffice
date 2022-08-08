using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.InsertPageBreak2
{
    public class InsertPageBreak2Controller : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public InsertPageBreak2Controller(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            //在文档末尾插入分页符，并且在新的页中创建新的数据区域插入另一篇文档
            PageOfficeNetCore.WordWriter.WordDocument wordDocument = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion mydr1 = wordDocument.CreateDataRegion("PO_first", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "[end]");
            mydr1.SelectEnd();
            wordDocument.InsertPageBreak();//插入分页符

            PageOfficeNetCore.WordWriter.DataRegion mydr2 = wordDocument.CreateDataRegion("PO_second", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "[end]");
            mydr2.Value = "[word]doc/test2.doc[/word]";

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.SetWriter(wordDocument);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test1.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/InsertPageBreak2/doc/" + "test3.doc");
            fs.CustomSaveResult = "ok";
            return  fs.Close();
            
        }
    }
}