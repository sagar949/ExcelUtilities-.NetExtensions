using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Utilities.Extensions;
namespace NetConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var persons = new List<Person>
            {
                new Person {Age = 24,Name="Sam"},
                new Person {Age = 25,Name="Sagar"},
                new Person {Age = 26,Name="Vidya"},
                new Person {Age = 27,Name="Test"},
                new Person {Age = 28,Name="Anderson"},
            };
            persons.SaveToExcel(fileName: "First");
            persons.SaveToExcel(fileName: "Second", removeColumns: new List<string> { "Age" });
            Console.WriteLine("Done!!!!!!!!!!!!!!!!!");
        }
    }

    class Person
    {
        public int Age { get; set; }

        public string Name { get; set; }
    }
}
