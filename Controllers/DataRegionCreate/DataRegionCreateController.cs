using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DataRegionCreate
{
    public class DataRegionCreateController : Controller
    {
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument wordDoc = new PageOfficeNetCore.WordWriter.WordDocument();
            //CreateDataRegion方法的三个参数分别代表：将要新建数据区域处的标签的名称、DataRegion的插入位置、与将要创建的DataRegion相关联的书签名称
            //若打开的Word文档中尚无书签或者想在Word文档的开头新建数据区域，则第三那个参数使用“[home]”若想在结尾处新建使用“[end]”
            PageOfficeNetCore.WordWriter.DataRegion dataRegion1 = wordDoc.CreateDataRegion("createDataRegion1", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "[home]");
            //为创建的DataRegion赋值
            dataRegion1.Value = "新建DataRegion1\r\n";

            PageOfficeNetCore.WordWriter.DataRegion dataRegion2 = wordDoc.CreateDataRegion("createDataRegion2", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "createDataRegion1");
            dataRegion2.Value = "新建DataRegion2\r\n";

            pageofficeCtrl.SetWriter(wordDoc);
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
    }
}