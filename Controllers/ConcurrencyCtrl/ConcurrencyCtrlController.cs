using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ConcurrencyCtrl
{
    public class ConcurrencyCtrlController : Controller
    {

        private readonly IWebHostEnvironment _webHostEnvironment;

        public ConcurrencyCtrlController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Word()
        {

            String userName = "somebody";
            String userId = Request.Query["userid"];
            if (userId.Equals("1"))
            {
                userName = "张三";
            }
            else
            {
                userName = "李四";
            }

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //设置并发控制时间
            pageofficeCtrl.TimeSlice = 20; // 设置并发控制时间, 单位:分钟
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docRevisionOnly, userName);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.userName = userName;
            return View();
        }
        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/ConcurrencyCtrl/doc/" + fs.FileName);
            return fs.Close();
            
        }

    }
}