using OfficeOpenXml;
using StudentExcelUpload.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace StudentExcelUpload.Controllers
{
    public class StudentsController : Controller
    {
        private readonly StudentContext db = new StudentContext();

        public async Task<ActionResult> Index()
        {
            var students = await db.Students.ToListAsync();
            return View(students);
        }

        [HttpPost]
        public async Task<ActionResult> Index(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                using (var package = new ExcelPackage(file.InputStream))
                {
                    var worksheet = package.Workbook.Worksheets.First();
                    var rowCount = worksheet.Dimension.Rows;

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

            var studentsList = await db.Students.ToListAsync();
            return View(studentsList);
        }



        public async Task<ActionResult> ExportToExcel()
        {
            var students = await db.Students.ToListAsync();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Students");

                // Add headers
                worksheet.Cells[1, 1].Value = "S.No";
                worksheet.Cells[1, 2].Value = "Student Name";
                worksheet.Cells[1, 3].Value = "Last Name";
                worksheet.Cells[1, 4].Value = "Date of Birth";
                worksheet.Cells[1, 5].Value = "Department";
                worksheet.Cells[1, 6].Value = "Graduation Year";
                worksheet.Cells[1, 7].Value = "Backlogs";
                worksheet.Cells[1, 8].Value = "CGPA";

                // Add values
                for (int i = 0; i < students.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = students[i].SNo;
                    worksheet.Cells[i + 2, 2].Value = students[i].StudentName;
                    worksheet.Cells[i + 2, 3].Value = students[i].LastName;
                    worksheet.Cells[i + 2, 4].Value = students[i].DateOfBirth.ToString("yyyy-MM-dd");
                    worksheet.Cells[i + 2, 5].Value = students[i].Department;
                    worksheet.Cells[i + 2, 6].Value = students[i].GraduationYear;
                    worksheet.Cells[i + 2, 7].Value = students[i].Backlogs;
                    worksheet.Cells[i + 2, 8].Value = students[i].CGPA;
                }

                // Set the content type and attachment header
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;
                string excelName = $"StudentList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
        }
    }

}