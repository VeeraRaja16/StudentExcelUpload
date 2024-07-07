using OfficeOpenXml;
using StudentExcelUpload.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace StudentExcelUpload.Controllers
{
    [RoutePrefix("api/upload")]
    public class UploadController : ApiController
    {
        [HttpPost]
        [Route("uploadexcel")]
        public async Task<IHttpActionResult> UploadExcel()
        {
            if (!Request.Content.IsMimeMultipartContent())
                return BadRequest("Unsupported media type");

            var provider = new MultipartMemoryStreamProvider();
            await Request.Content.ReadAsMultipartAsync(provider);

            foreach (var file in provider.Contents)
            {
                var dataStream = await file.ReadAsStreamAsync();
                using (var package = new ExcelPackage(dataStream))
                {
                    var worksheet = package.Workbook.Worksheets.First();
                    var rowCount = worksheet.Dimension.Rows;

                    using (var db = new StudentContext())
                    {
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var student = new Student
                            {
                                SNo = int.Parse(worksheet.Cells[row, 1].Text),
                                StudentName = worksheet.Cells[row, 2].Text,
                                LastName = worksheet.Cells[row, 3].Text,
                                DateOfBirth = DateTime.Parse(worksheet.Cells[row, 4].Text),
                                Department = worksheet.Cells[row, 5].Text,
                                GraduationYear = int.Parse(worksheet.Cells[row, 6].Text),
                                Backlogs = worksheet.Cells[row, 7].Text,
                                CGPA = decimal.Parse(worksheet.Cells[row, 8].Text)
                            };
                    db.Students.Add(student);
                }
                await db.SaveChangesAsync();
            }
        }
    }
        return Ok("File uploaded and data saved successfully.");
}
}

}
