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
using DataDictionary.Models.ResponseModels;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;

namespace DataDictionary.Controllers
{
  
    
    public class FileUploadController : Controller

    {
        private IHostingEnvironment _environment;
        private EahouseContext _context;
        private IDataMixingContract _datamixing;

        public FileUploadController(EahouseContext context, IDataMixingContract datamixing, IHostingEnvironment environment)
        {
            _context = context;
            _datamixing = datamixing;
            _environment = environment;
        }

[HttpPost]
        public async Task<ActionResult> UploadExcelFiles([FromForm]IFormFile files_, [FromForm]string version)
        {
            try
            {

                var dmm = new DataMixingExcelDocument();
                var data = 0;
                var dataDoc = (from s in _context.DataMixingExcelDocument



                               select new DataMixingExcelDocument
                               {
                                   DataMixingExcelDocumentId = s.DataMixingExcelDocumentId,
                                   FileName = s.FileName,
                                   ContentType = s.ContentType,
                                   FileSize = s.FileSize,
                                   DataMixingMasterId = s.DataMixingMasterId


                               }).ToList();

                foreach (var i in dataDoc)
                {


                    var item = _context.DataMixingExcelDocument.FirstOrDefault(o => o.DataMixingExcelDocumentId == i.DataMixingExcelDocumentId);
                    if (item == null)
                    {

                        data = 0;
                    }
                    else
                    {
                        data = (from MaxValue in _context.DataMixingExcelDocument
                                select MaxValue.DataMixingExcelDocumentId
                           ).Max();
                    }
                }






                var roothPath = Path.Combine(_environment.ContentRootPath, "excelLinkFile");
                if (!Directory.Exists(roothPath))
                    Directory.CreateDirectory(roothPath);


                var filePath = Path.Combine(roothPath, files_.FileName.Replace("ğ", "g").Replace("ı", "i").Replace("ş", "s").Replace("Ğ", "G").Replace("İ", "I").Replace("Ş", "S"));
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    var item = _context.DataMixingExcelDocument.FirstOrDefault(o => o.DataMixingMasterId == int.Parse(version));
                    if (item == null)
                    {
                        var document = new DataMixingExcelDocument
                        {
                            DataMixingExcelDocumentId = data + 1,
                            FileName = files_.FileName,
                            ContentType = files_.ContentType,
                            FileSize = files_.Length,
                            DataMixingMasterId = int.Parse(version),
                        };
                        await files_.CopyToAsync(stream);
                        _context.DataMixingExcelDocument.Add(document);
                        await _context.SaveChangesAsync();
                    }
                    else
                    {
                        item.DataMixingExcelDocumentId = item.DataMixingExcelDocumentId;
                        item.FileName = files_.FileName;
                        item.ContentType = files_.ContentType;
                        item.FileSize = files_.Length;
                        item.DataMixingMasterId = int.Parse(version);
                        await files_.CopyToAsync(stream);
                        _context.DataMixingExcelDocument.Update(item);
                        await _context.SaveChangesAsync();
                    }

                    var oldData = _context.DataMixingMaster.FirstOrDefault(x => x.DataMixingMasterId == int.Parse(version));
                    oldData.DataMixingMasterId = oldData.DataMixingMasterId;
                    oldData.DataMixingVersion = oldData.DataMixingVersion;
                    oldData.DataMixingVersionDescription = oldData.DataMixingVersionDescription;
                    oldData.IsApproved = oldData.IsApproved;
                    oldData.FromDate = DateTime.Now;
                    oldData.ExcelLink = (data + 1);
                    oldData.ApprovalMailLink = oldData.ApprovalMailLink;


                    _context.DataMixingMaster.Update(oldData);
                    await _context.SaveChangesAsync();




                }
                return Ok(data + 1);
            }
            catch (Exception e)
            {
                return NotFound(e);
            }



        }

}
}
       
