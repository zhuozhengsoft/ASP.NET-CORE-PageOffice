using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SaveFirstPageAsImg
{
    public class SaveFirstPageAsImgController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public SaveFirstPageAsImgController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("保存首页为图片", "SaveFirstAsImg()", 1);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;

            if (fs.FileExtName.Equals(".jpg"))
            {
                fs.SaveToFile(webRootPath + "/SaveFirstPageAsImg/images/" + fs.FileName);
            }
            else
            {
                fs.SaveToFile(webRootPath + "/SaveFirstPageAsImg/doc/" + fs.FileName);
            }

            fs.CustomSaveResult = "ok";
            return fs.Close();
            
            
        }
    }
}