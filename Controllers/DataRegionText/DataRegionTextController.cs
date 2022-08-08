using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using PageOfficeNetCore.WordWriter;

namespace NetCoreSamples5.Controllers.DataRegionText
{
    public class DataRegionTextController : Controller
    {

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion d1 = doc.OpenDataRegion("d1");
            d1.Font.Color = Color.Green;//设置数据区域文本字体颜色
            d1.Font.Name = "华文彩云";//设置数据区域文本字体样式
            d1.Font.Size = 16;//设置数据区域文本字体大小
            d1.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//设置数据区域文本对齐方式

            PageOfficeNetCore.WordWriter.DataRegion d2 = doc.OpenDataRegion("d2");
            d2.Font.Color = Color.MediumAquamarine;//设置数据区域文本字体颜色
            d2.Font.Name = "黑体";//设置数据区域文本字体样式
            d2.Font.Size = 14;//设置数据区域文本字体大小
            d2.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//设置数据区域文本对齐方式

            PageOfficeNetCore.WordWriter.DataRegion d3 = doc.OpenDataRegion("d3");
            d3.Font.Color = Color.Purple;//设置数据区域文本字体颜色
            d3.Font.Name = "华文行楷";//设置数据区域文本字体样式
            d3.Font.Size = 12;//设置数据区域文本字体大小
            d3.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;//设置数据区域文本对齐方式
            pageofficeCtrl.SetWriter(doc);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public IActionResult Word2()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}