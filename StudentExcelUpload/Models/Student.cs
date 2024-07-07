using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace StudentExcelUpload.Models
{
    public class Student
    {
        public int Id { get; set; }
        public int SNo { get; set; }
        public string StudentName { get; set; }
        public string LastName { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string Department { get; set; }
        public int GraduationYear { get; set; }
        public string Backlogs { get; set; }
        public decimal CGPA { get; set; }
    }

}