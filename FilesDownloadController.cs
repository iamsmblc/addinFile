using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using DevExtreme.AspNet.Data;
using Newtonsoft.Json;
using Microsoft.EntityFrameworkCore;
using DataDictionaryWebApi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Cors;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using System.Data;
using DataDictionary.Entities;
using Microsoft.AspNetCore.Http;
using System.IO;
using Microsoft.Extensions.Hosting;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.StaticFiles;

namespace DataDictionary.Controllers
{
  
    public class FilesDownloadController : Controller

    {
        private IHostingEnvironment _environment;
        private EahouseContext _context;
        private IDataMixingContract _datamixing;

        public FilesDownloadController (EahouseContext context, IDataMixingContract datamixing, IHostingEnvironment environment)
        {
            _context = context;
            _datamixing = datamixing;
            _environment = environment;
        }
    
 
    
     
     
        [HttpGet]
        public async Task<ActionResult> DownloadExcel(string versionName, string fileName)
        {
            try
            {
                var provider = new FileExtensionContentTypeProvider();
                var document = await _context.DataMixingExcelDocument.FirstOrDefaultAsync(o => o.FileName == fileName);
                if (document == null)
                    return NotFound();
                var file = Path.Combine(_environment.ContentRootPath, "excelLinkFile", document.FileName);
                string contentType;
                if (!provider.TryGetContentType(file, out contentType))
                {
                    contentType = "application/octet-stream";
                }
                byte[] fileBytes;
                if (System.IO.File.Exists(file))
                {

                    fileBytes = System.IO.File.ReadAllBytes(file);
                }
                else
                {
                    return NotFound();
                }
                MemoryStream ms = new MemoryStream(fileBytes);
                return new FileStreamResult(ms, contentType);


            }
            catch (Exception e)
            {
                return NotFound(e);
            }
        }



        }
}




