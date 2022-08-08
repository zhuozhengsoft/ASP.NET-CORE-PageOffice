using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordRibbonCtrl
{
    public class WordRibbonCtrlController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.RibbonBar.SetTabVisible("TabHome", true);//开始的Ribbon
            pageofficeCtrl.RibbonBar.SetTabVisible("TabInsert", false);//插入的Ribbon
            pageofficeCtrl.RibbonBar.SetTabVisible("TabPageLayoutWord", false);//页面布局的Ribbon
            pageofficeCtrl.RibbonBar.SetTabVisible("TabReferences", false);//引用的Ribbon
            pageofficeCtrl.RibbonBar.SetTabVisible("TabMailings", false);//邮件的Ribbon
            pageofficeCtrl.RibbonBar.SetTabVisible("TabView", false);//视图的Ribbon
            pageofficeCtrl.RibbonBar.SetTabVisible("TabReviewWord", false);//审阅的Ribbon

            pageofficeCtrl.RibbonBar.SetSharedVisible("FileSave", false);//office自带的保存按钮
            pageofficeCtrl.RibbonBar.SetGroupVisible("GroupClipboard", false);//开始中的剪切板组
            pageofficeCtrl.AddCustomToolButton("保存", "SaveFile()", 1);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}