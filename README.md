# ASP.NET-CORE-PageOffice

### 一、简介

​      ASP.NET-CORE-PageOffice项目演示了在ASP.NET CORE框架下如何使用PageOffice产品，此项目演示了PageOffice产品近90%的功能，是一个demo项目。

### 二、项目环境要求

​    Visual Studio2019 及以上版本。

### 三、项目运行步骤

1. 使用git clone或者直接下载项目压缩包到本地并解压缩。
2. 双击运行ASP.NET-CORE-PageOffice目录下的NetCoreSamples5.sln，然后运行示例并访问/index页面查看示例效果。

### 四、PageOffice序列号

​     PageOfficeV5.0标准版试用序列号：I2BFU-MQ89-M4ZZ-ZWY7K           
​     PageOfficeV5.0专业版试用序列号：DJMTF-HYK4-BDQ3-2MBUC

### 五、集成PageOffice到您的项目中的关键步骤

1. 在您的web项目的“依赖项-包-管理NuGet程序包”中搜索到“Zhuozhengsoft.PageOffice"程序后安装最新的版本。 
2. 拷贝“ASP.NET-CORE-PageOffice/Controllers”目录下的PageOfficeController.cs文件到您项目的Controllers文件夹下。
3. 在您项目的wwwroot文件夹下新建lic文件夹，此文件夹用来存放PageOffice的授权文件。
4. 对PageOffice编程控制：

   (1) 后台代码，在需要调用PageOffice的Controller中添加如下代码(详细代码请参考ASP.NET-CORE-PageOffice/Controllers/SimpleWord /SimpleWordController.cs文件)。

```c#
public IActionResult Word()
        {
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "/POserver";
            //添加自定义按钮
            pageofficeCtrl.AddCustomToolButton("保存", "Save()", 1);
            pageofficeCtrl.AddCustomToolButton("另存到本地", "SaveAs", 12);
            pageofficeCtrl.AddCustomToolButton("打印设置", "PrintSet", 0);
            pageofficeCtrl.AddCustomToolButton("打印", "PrintFile", 6);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            pageofficeCtrl.AddCustomToolButton("-", "", 0);
            pageofficeCtrl.AddCustomToolButton("关闭", "Close", 21);
            pageofficeCtrl.FileTitle="test";//设置另存为到本地的文件名称
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("doc/test.doc",PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }
```

  (2) 前台页面代码：(详细代码请参考ASP.NET-CORE-PageOffice/Views/SimpleWord/Word.cshtml文件)：

```c#
@Html.Raw(ViewBag.POCtrl)
```

5. 如果要使用“PageOffice浏览器”方式打开文件， 那么需要调用 javascript 方法

“POBrowser.openWindowModeless”的页面一定要引用下面的 js 文件：

`<script type="text/javascript" src="/pageoffice.js"></script>`

> 【注意】：pageoffice.js 文件已经在第2步拷贝的PageOfficeController.cs文件中配置好了引用路径，默认pageoffice.js文件配置到了当前项目的根目录下，所以需要调用POBrowser.openWindowModeless的页面直接引用当前项目根目录下的这个 js 即可，无需拷贝 pageoffice.js 文件到自己的Web项目目录下。

### 六、 PageOffice开发帮助

​     1 .[JS API文档](https://www.zhuozhengsoft.com/help/js3/index.html)  

​     2 .[PageOffice从入门到精通](https://www.kancloud.cn/pageoffice_course_group/pageoffice_course/646953)

​     技术支持：https://www.zhuozhengsoft.com/Technical/

### 七、联系我们

​   卓正官网：[https://www.zhuozhengsoft.com](https://www.zhuozhengsoft.com)

​   联系电话：400-6600-770  

   QQ: 800038353