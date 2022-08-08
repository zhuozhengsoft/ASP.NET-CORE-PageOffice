using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.TaoHong
{
    public class TaoHongController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        private readonly IWebHostEnvironment _webHostEnvironment;

        public TaoHongController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult edit()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult taoHong()
        {


            String mb = Request.Query["mb"];
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            String fileName = "test.doc";
            if (mb != null && mb.Trim() != "")
            {
                string webRootPath = _webHostEnvironment.WebRootPath;

                fileName = "zhengshi.doc";

                System.IO.File.Copy(webRootPath + "/TaoHong/doc/" + mb,
               webRootPath + "/TaoHong/doc/" + fileName, true);
                //给正式发文的文件填充数据
                PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
                PageOfficeNetCore.WordWriter.DataRegion copies = doc.OpenDataRegion("PO_Copies");
                copies.Value = "6";
                PageOfficeNetCore.WordWriter.DataRegion docNum = doc.OpenDataRegion("PO_DocNum");
                docNum.Value = "001";
                PageOfficeNetCore.WordWriter.DataRegion issueDate = doc.OpenDataRegion("PO_IssueDate");
                issueDate.Value = "2013-5-30";
                PageOfficeNetCore.WordWriter.DataRegion issueDept = doc.OpenDataRegion("PO_IssueDept");
                issueDept.Value = "开发部";
                PageOfficeNetCore.WordWriter.DataRegion sTextS = doc.OpenDataRegion("PO_STextS");
                sTextS.Value = "[word]doc/test.doc[/word]"; // 把正文文件插入到正式文件中
                PageOfficeNetCore.WordWriter.DataRegion sTitle = doc.OpenDataRegion("PO_sTitle");
                sTitle.Value = "北京某公司文件";
                PageOfficeNetCore.WordWriter.DataRegion topicWords = doc.OpenDataRegion("PO_TopicWords");
                topicWords.Value = "Pageoffice、 套红";
                pageofficeCtrl.SetWriter(doc);
            }

            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            pageofficeCtrl.WebOpen("doc/" + fileName, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");

            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
        public IActionResult readOnly()
        {
            string fileName = "zhengshi.doc";//正式发文文件
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("另存到本地", "ShowDialog(0)", 5);
            pageofficeCtrl.AddCustomToolButton("页面设置", "ShowDialog(1)", 0);
            pageofficeCtrl.AddCustomToolButton("打印", "ShowDialog(2)", 6);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen()", 4);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/" + fileName, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/TaoHong/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}