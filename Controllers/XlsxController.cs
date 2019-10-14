using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using CsvConvert.Helpers;
using CsvConvert.Services;
using Microsoft.AspNetCore.Mvc;

namespace CsvConvert.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class XlsxController : ControllerBase
    {
        private IXlsService xlsService;
        public XlsxController(IXlsService xlsService)
        {
            this.xlsService = xlsService;
        }

        [HttpPost("convert_one")]
        public async Task<ActionResult> Post()
        {
            try
            {
                var file = Request.Form.Files[0];
                var folderName = Path.Combine("Resources", "Xlsx");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);

                if (file.Length > 0)
                {
                    var fileName = Guid.NewGuid().ToString() + "_" + ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
                    var fullPath = Path.Combine(pathToSave, fileName);

                    var dbPath = Path.Combine(folderName, fileName);

                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                    var resultFile = await xlsService.ToXlsx(dbPath);

                    var memory = new MemoryStream();
                    using (var stream = new FileStream(resultFile, FileMode.Open))
                    {
                        await stream.CopyToAsync(memory);
                    }
                    memory.Position = 0;
                    return File(memory, "application/octet-stream", Path.GetFileName(resultFile));
                }
                else
                {
                    //SimpleLogger.Log("inside3");
                    var tooSmall = "file is too small";
                    return BadRequest(new { tooSmall });
                }
            }
            catch
            {
                return StatusCode(500, "Internal server error");
            }
        }
    }
}