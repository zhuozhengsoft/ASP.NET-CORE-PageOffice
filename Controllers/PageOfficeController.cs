using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;

namespace NetCoreSamples5.Controllers
{
    public class PageOfficeController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;

        public PageOfficeController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        [Route("/POserver")]
        [Route("/pageoffice.js")]
        [Route("/pobstyle.css")]
        [Route("/posetup.exe")]
        [Route("/sealsetup.exe")]
        public ActionResult POServer()
        {
            PageOfficeNetCore.POServer.Server poServer = new PageOfficeNetCore.POServer.Server(Request, Response);
            poServer.LicenseFilePath = _webHostEnvironment.ContentRootPath + "/wwwroot/lic/";
            return poServer.Run();
            
        }

    }
}