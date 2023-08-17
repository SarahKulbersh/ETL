using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TRY
{
    public class Employee
    {
        public double Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public double Age { get; set; }


        public Employee(double id, string firstName, string lastName, double age)
        {
            Id = id;
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }

    }
}
