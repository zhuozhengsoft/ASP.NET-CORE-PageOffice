using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.SendParameters
{
    public class SendParametersController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public SendParametersController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            pageofficeCtrl.Caption = "演示：向保存页面传递参数，更新人员信息";
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("全屏", "SetFullScreen()", 4);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc?id=1";//传递查询参数
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }


        public async Task<ActionResult> SaveDoc()
        {
            int id = 0;
            string userName = "";
            int age = 0;
            string sex = "";
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            //await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/SendParameters/doc/" + fs.FileName);

            //获取通过Url传递过来的值
            string message = Request.Query["id"];

            if (message != null && message.Trim().Length > 0)
                id = int.Parse(message.Trim());

            //获取通过网页标签控件传递过来的参数值，注意fs.GetFormField("HTML标签的name名称")方法中的参数名是指标签的Id

            //获取通过文本框<input type="text" />标签传递过来的值
            if (fs.GetFormField("userName") != null && fs.GetFormField("userName").Trim().Length > 0)
            {
                userName = fs.GetFormField("userName");
            }

            //获取通过隐藏域传递过来的值
            if (fs.GetFormField("age") != null && fs.GetFormField("age").Trim().Length > 0)
            {
                age = int.Parse(fs.GetFormField("age"));
            }

            //获取通过<select>标签传递过来的值
            if (fs.GetFormField("selSex") != null && fs.GetFormField("selSex").Trim().Length > 0)
            {
                sex = fs.GetFormField("selSex");
            }

            fs.ShowPage(300, 200,this); // 显示一下SaveFile.aspx获取到的所有参数的值


            string content = "";
            content += "传递的参数为：<br />";
            content += " userName:" + userName + "<br />";
            content += " id:" + id + "<br />";
            content += " age:" + age + "<br />";
            content += " sex:" + sex + "<br />";

            //await Response.Body.WriteAsync(Encoding.GetEncoding("gbk").GetBytes(content));
            fs.Close();
            ViewBag.content = content;
            return View();
        }
    }
}