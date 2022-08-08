using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SetDrByUserWord
{
    public class SetDrByUserWordController : Controller
    {

        private readonly IWebHostEnvironment _webHostEnvironment;

        public SetDrByUserWordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Word()
        {

            string user = "";
            string userName = Request.Form["userName"];

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //***************************卓正PageOffice组件的使用********************************
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();

            PageOfficeNetCore.WordWriter.DataRegion dA1 = doc.OpenDataRegion("PO_A_pro1");
            PageOfficeNetCore.WordWriter.DataRegion dA2 = doc.OpenDataRegion("PO_A_pro2");
            PageOfficeNetCore.WordWriter.DataRegion dB1 = doc.OpenDataRegion("PO_B_pro1");
            PageOfficeNetCore.WordWriter.DataRegion dB2 = doc.OpenDataRegion("PO_B_pro2");

            //根据登录用户名设置数据区域可编辑性
            //A部门经理登录后
            if (userName.Equals("zhangsan"))
            {
                dA1.Editing = true;
                dA2.Editing = true;
                dB1.Editing = false;
                dB2.Editing = false;
                user = "A部门经理";
            }
            //B部门经理登录后
            else
            {
                dB1.Editing = true;
                dB2.Editing = true;
                dA1.Editing = false;
                dA2.Editing = false;
                user = "B部门经理";
            }

            pageofficeCtrl.SetWriter(doc);
            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, user);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.user = user;
            return View();
        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SetDrByUserWord/doc/" + fs.FileName);
            return fs.Close();
            
        }
    }
}