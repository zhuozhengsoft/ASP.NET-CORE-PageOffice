using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExcelRibbonCtrl
{
    public class ExcelRibbonCtrlController : Controller
    {
        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.RibbonBar.SetTabVisible("TabHome", true);//开始
            pageofficeCtrl.RibbonBar.SetTabVisible("TabFormulas", false);//公式
            pageofficeCtrl.RibbonBar.SetTabVisible("TabInsert", false);//插入
            pageofficeCtrl.RibbonBar.SetTabVisible("TabData", false);//数据
            pageofficeCtrl.RibbonBar.SetTabVisible("TabReview", false);//审阅
            pageofficeCtrl.RibbonBar.SetTabVisible("TabView", false);//视图
            pageofficeCtrl.RibbonBar.SetTabVisible("TabPageLayoutExcel", false);//页面布局

            pageofficeCtrl.RibbonBar.SetSharedVisible("FileSave", false);//office自带的保存按钮

            //分组
            pageofficeCtrl.RibbonBar.SetGroupVisible("GroupClipboard", false);//剪贴板
            pageofficeCtrl.AddCustomToolButton("保存", "SaveFile()", 1);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}