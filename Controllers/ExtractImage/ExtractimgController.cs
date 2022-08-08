using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using System.Web;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExtractImage
{
    public class ExtractImageController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public ExtractImageController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义的保存按钮
            pageofficeCtrl.AddCustomToolButton("保存图片", "Save", 1);
            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion dataRegion1 = wordDoc.OpenDataRegion("PO_image");
            dataRegion1.Editing = true;//放图片的数据区域是可以编辑的，其它部分不可编辑
            pageofficeCtrl.SetWriter(wordDoc);

            //设置保存页面
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
            PageOfficeNetCore.WordReader.DataRegion dataRegion1 = doc.OpenDataRegion("PO_image");
            //将提取的图片保存到服务器上，图片的名称为:a.jpg

            string webRootPath = _webHostEnvironment.WebRootPath;
            dataRegion1.OpenShape(1).SaveAsJPG(webRootPath + "/ExtractImage/doc/a.jpg");
            //注册编码提供程序
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


            //UrlEncoder
            doc.CustomSaveResult = HttpUtility.UrlEncode("保存成功,文件保存到：" + "wwwroot/ExtractImage/doc/a.jpg");
            return doc.Close();
            
        }
    }
}