using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordTextBox
{
    public class WordTextBoxController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();

            wordDoc.OpenDataRegion("PO_company").Value = "北京幻想科技有限公司";
            wordDoc.OpenDataRegion("PO_logo").Value = "[image]doc/logo.gif[/image]";
            wordDoc.OpenDataRegion("PO_dr1").Value = "左边的文本:xxxx";

            pageofficeCtrl.SetWriter(wordDoc);// 注意不要忘记此代码，如果缺少此句代码，不会赋值成功。
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}