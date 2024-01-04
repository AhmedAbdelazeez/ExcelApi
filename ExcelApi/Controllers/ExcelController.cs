using ExcelApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly Aplicationdbcontext _context;

        public ExcelController(Aplicationdbcontext context)
        {
            _context = context;
        }




        [HttpPost("Import")]
        public async Task<IActionResult> import(IFormFile file)
        {
            var list = new List<RequestNumber>();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    if (worksheet != null)
                    {
                        var rowcount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowcount; row++)
                        {
                            var requestNumberCell = worksheet.Cells[row, 1];
                            var areaCell = worksheet.Cells[row, 2];

                            // Check if cells exist before accessing their values
                            if (requestNumberCell.Value != null && areaCell.Value != null)
                            {
                                list.Add(new RequestNumber
                                {
                                    requestNumber = requestNumberCell.Value.ToString().Trim(),
                                    Area = areaCell.Value.ToString().Trim()
                                });
                            }
                            foreach (var item in list)
                            {
                                var requestToUpdate = _context.tests.FirstOrDefault(r => r.name == item.requestNumber);

                                if (requestToUpdate != null)
                                {
                                    // Update the 'Area' property
                                    requestToUpdate.area = decimal.Parse(item.Area);

                                    // Save changes to the database
                                    await _context.SaveChangesAsync();
                                }
                                // Handle the case where the request number is not found if needed.
                            }
                        }
                    }
                    else
                    {
                        // Handle the case where the worksheet is null
                        return BadRequest("Worksheet is null.");
                    }
                }
            }
            return Ok("Upadted Successsfuly pro.");
            // The rest of your code remains unchanged...
        }

        //[HttpPost("Imort")]
        //public async Task<List<RequestNumber>> Imort(IFormFile file)
        //{
        //    var list = new List<RequestNumber>();
        //    using (var stream = new MemoryStream())
        //    {
        //        await file.CopyToAsync(stream);
        //        using (var package = new ExcelPackage(stream))
        //        {
        //            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        //            var rowcount = worksheet.Dimension.Rows;
        //            for (int row = 2; row <= rowcount; row++)
        //            {
        //                list.Add(new RequestNumber
        //                {
        //                    requestNumber = worksheet.Cells[row, 1].Value.ToString().Trim()
        //                });
        //            }
        //        }
        //    }
        //    return list;
        //}
    }

}
