using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.ExcelDisableRight
{
    public class ExcelDisableRightController : Controller
    {

        public IActionResult Excel()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            workBook.DisableSheetRightClick = true;//禁止当前工作表的鼠标右键                                                   //workBook.DisableSheetDoubleClick = true;//禁止当前工作表的鼠标双击
            pageofficeCtrl.SetWriter(workBook);

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.xls", PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}