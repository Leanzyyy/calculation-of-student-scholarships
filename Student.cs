using System.Collections.Generic;

namespace PRAK1
{
    public class Student
    {
        public string Name { get; set; } = "";
        public bool IsLargeFamily { get; set; }
        public List<int> Grades { get; set; } = new List<int>();
    }
}