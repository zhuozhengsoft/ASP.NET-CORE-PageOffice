using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace NetCoreSamples5.Controllers.WordParagraph
{
    public class WordParagraphController : Controller
    {

        public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";

            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();

            //设置内容标题

            //创建DataRegion对象，PO_title为自动添加的书签名称,书签名称需以“PO_”为前缀，切书签名称不能重复
            //三个参数分别为要新插入书签的名称、新书签的插入位置、相关联的书签名称（“[home]”代表Word文档的第一个位置）
            PageOfficeNetCore.WordWriter.DataRegion title = doc.CreateDataRegion("PO_title",
                PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "[home]");
            //给DataRegion对象赋值
            title.Value = "C#中Socket多线程编程实例\n";
            //设置字体：粗细、大小、字体名称、是否是斜体
            title.Font.Bold = true;
            title.Font.Size = 20;
            title.Font.Name = "黑体";
            title.Font.Italic = false;
            //定义段落对象
            PageOfficeNetCore.WordWriter.ParagraphFormat titlePara = title.ParagraphFormat;
            //设置段落对齐方式
            titlePara.Alignment = PageOfficeNetCore.WordWriter.WdParagraphAlignment.wdAlignParagraphCenter;
            //设置段落行间距
            titlePara.LineSpacingRule = PageOfficeNetCore.WordWriter.WdLineSpacing.wdLineSpaceMultiple;

            //设置内容
            //第一段
            //创建DataRegion对象，PO_body为自动添加的书签名称
            PageOfficeNetCore.WordWriter.DataRegion body = doc.CreateDataRegion("PO_body", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "PO_title");
            //设置字体：粗细、是否是斜体、大小、字体名称、字体颜色
            body.Font.Bold = false;
            body.Font.Italic = true;
            body.Font.Size = 10;
            //设置中文字体名称
            body.Font.Name = "楷体";
            //设置英文字体名称
            body.Font.NameAscii = "Times New Roman";
            body.Font.Color = Color.Red;
            //给DataRegion对象赋值
            body.Value = "是微软随着VS.net新推出的一门语言。它作为一门新兴的语言，有着C++的强健，又有着VB等的RAD特性。而且，微软推出C#主要的目的是为了对抗Sun公司的Java。大家都知道Java语言的强大功能，尤其在网络编程方面。于是，C#在网络编程方面也自然不甘落后于人。本文就向大家介绍一下C#下实现套接字（Sockets）编程的一些基本知识，以期能使大家对此有个大致了解。首先，我向大家介绍一下套接字的概念。\n";
            //创建ParagraphFormat对象
            PageOfficeNetCore.WordWriter.ParagraphFormat bodyPara = body.ParagraphFormat;
            //设置段落的行间距、对齐方式、首行缩进
            bodyPara.LineSpacingRule = PageOfficeNetCore.WordWriter.WdLineSpacing.wdLineSpaceAtLeast;
            bodyPara.Alignment = PageOfficeNetCore.WordWriter.WdParagraphAlignment.wdAlignParagraphLeft;
            bodyPara.FirstLineIndent = 21;

            //第二段
            PageOfficeNetCore.WordWriter.DataRegion body2 = doc.CreateDataRegion("PO_body2", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "PO_body");
            body2.Font.Bold = false;
            body2.Font.Size = 12;
            body2.Font.Name = "黑体";
            body2.Value = "套接字是通信的基石，是支持TCP/IP协议的网络通信的基本操作单元。可以将套接字看作不同主机间的进程进行双向通信的端点，它构成了单个主机内及整个网络间的编程界面。套接字存在于通信域中，通信域是为了处理一般的线程通过套接字通信而引进的一种抽象概念。套接字通常和同一个域中的套接字交换数据（数据交换也可能穿越域的界限，但这时一定要执行某种解释程序）。各种进程使用这个相同的域互相之间用Internet协议簇来进行通信。\n";
            //body2.Value ="[image]../images/logo.jpg[/image]";
            PageOfficeNetCore.WordWriter.ParagraphFormat bodyPara2 = body2.ParagraphFormat;
            bodyPara2.LineSpacingRule = PageOfficeNetCore.WordWriter.WdLineSpacing.wdLineSpace1pt5;
            bodyPara2.Alignment = PageOfficeNetCore.WordWriter.WdParagraphAlignment.wdAlignParagraphLeft;
            bodyPara2.FirstLineIndent = 21;

            //第三段
            PageOfficeNetCore.WordWriter.DataRegion body3 = doc.CreateDataRegion("PO_body3", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "PO_body2");
            body3.Font.Bold = false;
            body3.Font.Color = Color.FromArgb(0, 128, 128);
            body3.Font.Size = 14;
            body3.Font.Name = "华文彩云";
            body3.Value = "套接字可以根据通信性质分类，这种性质对于用户是可见的。应用程序一般仅在同一类的套接字间进行通信。不过只要底层的通信协议允许，不同类型的套接字间也照样可以通信。套接字有两种不同的类型：流套接字和数据报套接字。\n";
            PageOfficeNetCore.WordWriter.ParagraphFormat bodyPara3 = body3.ParagraphFormat;
            bodyPara3.LineSpacingRule = PageOfficeNetCore.WordWriter.WdLineSpacing.wdLineSpaceDouble;
            bodyPara3.Alignment = PageOfficeNetCore.WordWriter.WdParagraphAlignment.wdAlignParagraphLeft;
            bodyPara3.FirstLineIndent = 21;

            PageOfficeNetCore.WordWriter.DataRegion body4 = doc.CreateDataRegion("PO_body4", PageOfficeNetCore.WordWriter.DataRegionInsertType.After, "PO_body3");
            body4.Value = "[image]doc/logo.png[/image]";
            //body4.Value = "[word]doc/1.doc[/word]";//还可嵌入其他Word文件
            PageOfficeNetCore.WordWriter.ParagraphFormat bodyPara4 = body4.ParagraphFormat;
            bodyPara4.Alignment = PageOfficeNetCore.WordWriter.WdParagraphAlignment.wdAlignParagraphCenter;

            //PageOffice组件的使用

            //隐藏自定义工具栏
            pageofficeCtrl.CustomToolbar = false;
            pageofficeCtrl.SetWriter(doc);
            pageofficeCtrl.JsFunction_AfterDocumentSaved = "SaveOK()";

            //打开Word文档
            pageofficeCtrl.WebOpen("doc/template.doc", PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

    }
}