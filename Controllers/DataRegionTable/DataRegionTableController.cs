using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.DataRegionTable
{
    public class DataRegionTableController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public DataRegionTableController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            //打开数据区域
            PageOfficeNetCore.WordWriter.DataRegion dTable = doc.OpenDataRegion("PO_table");
            //设置数据区域可编辑性
            dTable.Editing = true;
            //打开数据区域中的表格，OpenTable(index)方法中的index为word文档中表格的下标，从1开始
            PageOfficeNetCore.WordWriter.Table table1 = doc.OpenDataRegion("PO_Table").OpenTable(1);
            // 给表头单元格赋值
            table1.OpenCellRC(1, 2).Value = "产品1";
            table1.OpenCellRC(1, 3).Value = "产品2";
            table1.OpenCellRC(2, 1).Value = "A部门";
            table1.OpenCellRC(3, 1).Value = "B部门";

            pageofficeCtrl.SetWriter(doc);

            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);

            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveData";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

        public async Task<ActionResult> SaveData()

        {

            //-----------  PageOffice 服务器端编程开始  -------------------//
            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);

            await doc.LoadAsync();
            PageOfficeNetCore.WordReader.DataRegion dataReg = doc.OpenDataRegion("PO_table");
            PageOfficeNetCore.WordReader.Table table = dataReg.OpenTable(1);
            //输出提交的table中的数据
            //Response.Write("表格中的各个单元的格数据为：<br/><br/>");
            StringBuilder dataStr = new StringBuilder();
            for (int i = 1; i <= table.RowsCount; i++)
            {
                dataStr.Append("<div style='width:220px;'>");
                for (int j = 1; j <= table.ColumnsCount; j++)
                {
                    dataStr.Append("<div style='float:left;width:70px;border:1px solid red;'>" + table.OpenCellRC(i, j).Value + "</div>");
                }
                dataStr.Append("</div>");
            }
            //Response.Write(dataStr.ToString());
            //向客户端显示提交的数据

            // await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(dataStr.ToString()));

            doc.ShowPage(300, 300,this);
            doc.Close();
            ViewBag.dataStr = dataStr;
            return View();
        }
    }
}