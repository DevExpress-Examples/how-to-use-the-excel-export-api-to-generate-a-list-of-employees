using System.Collections.Generic;
using System.Linq;

namespace XLExportExample {
    class EmployeeData {
        
        public EmployeeData(int id, string name, double salary, double bonus, string department) 
        {
            Id = id;
            Name = name;
            Salary = salary;
            Bonus = bonus;
            Department = department;
        }

        public int Id { get; private set; }
        public string Name { get; private set; }
        public double Salary { get; private set; }
        public double Bonus { get; private set; }
        public string Department { get; private set; }
    }

    static class EmployeesRepository {
        static string[] departments = new string[] { "Accounting", "Logistics", "IT", "Management", "Manufacturing", "Marketing" };

        public static List<EmployeeData> CreateEmployees() {
            List<EmployeeData> result = new List<EmployeeData>();
            result.Add(new EmployeeData(10115, "Augusta Delono", 1100.0, 50.0, "Accounting"));
            result.Add(new EmployeeData(10501, "Berry Dafoe", 1650.0, 150.0, "IT"));
            result.Add(new EmployeeData(10709, "Chris Cadwell", 2000.0, 180.0, "Management"));
            result.Add(new EmployeeData(10356, "Esta Mangold", 1400.0, 75.0, "Logistics"));
            result.Add(new EmployeeData(10401, "Frank Diamond", 1750.0, 100.0, "Marketing"));
            result.Add(new EmployeeData(10202, "Liam Bell", 1200.0, 80.0, "Manufacturing"));
            result.Add(new EmployeeData(10205, "Simon Newman", 1250.0, 80.0, "Manufacturing"));
            result.Add(new EmployeeData(10403, "Wendy Underwood", 1100.0, 50.0, "Marketing"));
            return result;
        }

        public static List<string> CreateDepartments() {
            return departments.ToList();
        }
    }
}
