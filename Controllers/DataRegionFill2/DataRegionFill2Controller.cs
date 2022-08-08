using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DataRegionFill2
{
    public class DataRegionFill2Controller : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();

            //打开数据区域，OpenDataRegion方法的参数代表Word文档中的书签名称
            PageOfficeNetCore.WordWriter.DataRegion dataRegion1 = wordDoc.OpenDataRegion("PO_userName");
            //为DataRegion赋值
            dataRegion1.Value = "张三";
            //设置字体样式
            dataRegion1.Font.Color = Color.Blue;
            dataRegion1.Font.Size = 24f;
            dataRegion1.Font.Name = "隶书";
            dataRegion1.Font.Bold = true;

            PageOfficeNetCore.WordWriter.DataRegion dataRegion2 = wordDoc.OpenDataRegion("PO_deptName");
            dataRegion2.Value = "人事部";
            dataRegion2.Font.Color = Color.Red;

            pageofficeCtrl.SetWriter(wordDoc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}