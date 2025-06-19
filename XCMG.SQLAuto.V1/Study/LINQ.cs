using MathNet.Numerics.Distributions;
using NPOI.SS.Formula.Functions;
using RekTec.Crm.Common.Helper;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace XCMG.SQLAuto.V1.Study
{
    public static class LINQ
    {
        record City(string Name, long Population);
        record Country(string Name, double Area, long Population, List<City> Cities);
        // record Product(string Name, string Category);

        static readonly City[] cities = [
            new City("Tokyo", 37_833_000),
            new City("Delhi", 30_290_000),
            new City("Shanghai", 27_110_000),
            new City("São Paulo", 22_043_000),
            new City("Mumbai", 20_412_000),
            new City("Beijing", 20_384_000),
            new City("Cairo", 18_772_000),
            new City("Dhaka", 17_598_000),
            new City("Osaka", 19_281_000),
            new City("New York-Newark", 18_604_000),
            new City("Karachi", 16_094_000),
            new City("Chongqing", 15_872_000),
            new City("Istanbul", 15_029_000),
            new City("Buenos Aires", 15_024_000),
            new City("Kolkata", 14_850_000),
            new City("Lagos", 14_368_000),
            new City("Kinshasa", 14_342_000),
            new City("Manila", 13_923_000),
            new City("Rio de Janeiro", 13_374_000),
            new City("Tianjin", 13_215_000)
        ];

        static readonly Country[] countries = [
            new Country ("Vatican City", 0.44, 526, [new City("Vatican City", 826)]),
            new Country ("Monaco", 2.02, 38_000, [new City("Monte Carlo", 38_000)]),
            new Country ("Nauru", 21, 10_900, [new City("Yaren", 1_100)]),
            new Country ("Tuvalu", 26, 11_600, [new City("Funafuti", 6_200)]),
            new Country ("San Marino", 61, 33_900, [new City("San Marino", 4_500)]),
            new Country ("Liechtenstein", 160, 38_000, [new City("Vaduz", 5_200)]),
            new Country ("Marshall Islands", 181, 58_000, [new City("Majuro", 28_000)]),
            new Country("Saint Kitts & Nevis", 261, 53_000, [new City("Basseterre",13_000)])
        ];

        record Product(string Name, int Number, string Category);
        record Category(string Name, int ID);

        record Student(string Name, string Last, string First, int Year, List<int> ExamScores);

        public static void SelectNewObject()
        {
            IEnumerable<City> cityQuery =
                from country in countries
                from city in country.Cities
                where city.Population > 10000
                select city;
            foreach (City city in cityQuery)
            {
                Console.WriteLine("FROM:" + city.Name);
            }

            var queryCountryGroups =
                from country in countries
                group country by country.Name[0] into countryGroup
                orderby countryGroup.Key descending
                select countryGroup;

            foreach (var queryCountryGroup in queryCountryGroups)
            {
                Console.WriteLine(queryCountryGroup.Key);
                foreach (var queryCountry in queryCountryGroup)
                {
                    Console.WriteLine("GROUP:" + JsonHelper.Serialize(queryCountry));
                }
            }

            //GROUP: [{ "Name":"Vatican City","Area":0.44,"Population":526,"Cities":[{ "Name":"Vatican City","Population":826}]}]
            //GROUP: [{ "Name":"Monaco","Area":2.02,"Population":38000,"Cities":[{ "Name":"Monte Carlo","Population":38000}]},{ "Name":"Marshall Islands","Area":181.0,"Population":58000,"Cities":[{ "Name":"Majuro","Population":28000}]}]
            //GROUP: [{ "Name":"Nauru","Area":21.0,"Population":10900,"Cities":[{ "Name":"Yaren","Population":1100}]}]
            //GROUP: [{ "Name":"Tuvalu","Area":26.0,"Population":11600,"Cities":[{ "Name":"Funafuti","Population":6200}]}]
            //GROUP: [{ "Name":"San Marino","Area":61.0,"Population":33900,"Cities":[{ "Name":"San Marino","Population":4500}]},{ "Name":"Saint Kitts & Nevis","Area":261.0,"Population":53000,"Cities":[{ "Name":"Basseterre","Population":13000}]}]
            //GROUP: [{ "Name":"Liechtenstein","Area":160.0,"Population":38000,"Cities":[{ "Name":"Vaduz","Population":5200}]}]

            var queryNameAndPop =
                from country in countries
                select new
                {
                    country.Name,
                    Pop = country.Population
                };


            var percentileQuery =
                from country in countries
                let percentile = (int)country.Population / 1_000
                group country by percentile into countryGroup
                where countryGroup.Key >= 20
                orderby countryGroup.Key
                select countryGroup;

            // grouping is an IGrouping<int, Country>
            foreach (var grouping in percentileQuery)
            {
                Console.WriteLine(grouping.Key);
                foreach (var country in grouping)
                {
                    Console.WriteLine("INTO:" + country.Name + ":" + country.Population);
                }
            }

            //33
            //INTO: San Marino:33900
            //38
            //INTO: Monaco: 38000
            //INTO: Liechtenstein: 38000
            //53
            //INTO: Saint Kitts &Nevis:53000
            //58
            //INTO: Marshall Islands:58000

            IEnumerable<City> queryCityPop =
                from city in cities
                where city.Population is < 15_000_000 and > 10_000_000
                select city;

            IEnumerable<Country> querySortedCountries =
                from country in countries
                orderby country.Area, country.Population descending
                select country;

            string[] categories = ["Orchestral", "Piano"];
            Product?[] products =
            [
                new Product("Trumpet", 1, "Orchestral"),
                new Product("Trombone", 1, "Orchestral"),
                new Product("French Horn", 1, "Orchestral"),
                null,
                new Product("Clarinet", 2, "Orchestral"),
                new Product("Flute", 2, "Orchestral"),
                null,
                new Product("Cymbal", 3, "Percussion"),
                new Product("Drum", 3, "Percussion")
            ];

            var categoryQuery =
                from cat in categories
                join prod in products on cat equals prod?.Category
                select new
                {
                    Category = cat,
                    Name = prod.Name
                };

            //var query =
            //    from o in db.Orders
            //    join e in db.Employees
            //    on o.EmployeeID equals (int?)e.EmployeeID // 可以为null
            //    select new { o.OrderID, e.FirstName };

            foreach (var category in categoryQuery)
            {
                Console.WriteLine("JOIN:" + JsonHelper.Serialize(category));
            }

            // let
            string[] names = ["Svetlana Omelchenko", "Claire O'Donnell", "SvenMortensen", "Cesar Garcia"];
            IEnumerable<string> queryFirstNames =
                from name in names
                let firstName = name.Split(' ')[0]
                select firstName;
            foreach (var s in queryFirstNames)
            {
                Console.Write(s + " ");
            }


            Student?[] students =
            [
                new Student ( "Alice", "Lance", "Tucker",2021, [90, 95, 100, 93] ),
                new Student ("Bob", "Terry", "Adams", 2021, [85, 80, 90, 84]),
                new Student ("Charlie", "Eugene", "Zabokritski", 2022, [88, 92, 96, 100]),
            ];

            var queryGroup =
                from student in students
                group student by student.Year into studentGroup
                select studentGroup;
            Console.WriteLine("StudentGroup:" + JsonHelper.Serialize(queryGroup));

            // [
            //      [
            //          { "Name":"Alice","Year":2021,"ExamScores":[90, 95, 100]},
            //          { "Name":"Bob","Year":2021,"ExamScores":[85, 80, 90]}
            //      ],
            //      [
            //          { "Name":"Charlie","Year":2022,"ExamScores":[88, 92, 96]}
            //      ]
            // ]

            var queryGroupMax =
                from student in students
                group student by student.Year into studentGroup
                select new
                {
                    Level = studentGroup.Key,
                    HighestScore = (
                        from student2 in studentGroup
                        select student2.ExamScores.Average()
                    ).Max()
                };

            string[] groupingQuery = ["carrots", "cabbage", "broccoli", "beans", "barley"];
            IEnumerable<IGrouping<char, string>> queryFoodGroups =
                from item in groupingQuery
                group item by item[0];
            Console.WriteLine("groupingQuery:" + JsonHelper.Serialize(queryFoodGroups));

            var studentQuery5 =
                from student in students
                let totalScore = student.ExamScores.Sum()
                where totalScore / 4 < student.ExamScores[0]
                select $"{student.Last}, {student.First}";


            static IEnumerable<int> GetData() => throw new InvalidOperationException();
            // DO THIS with a datasource that might
            // throw an exception.
            IEnumerable<int>? dataSource = null;
            try
            {
                dataSource = GetData();
            }
            catch (InvalidOperationException)
            {
                Console.WriteLine("Invalid operation");
            }

            if (dataSource is not null)
            {
                // If we get here, it is safe to proceed.
                var _query = from i in dataSource
                            select i * i;
                foreach (var i in _query)
                {
                    Console.WriteLine(i.ToString());
                }
            }

            //for (int i = 1; i < 10; i++)
            //{
            //    for (int j = 1; j <= i; j++)
            //    {
            //        if (j == 1) Console.WriteLine();
            //        Console.Write($"{j}*{i}={j * i}" + "\t");
            //    }
            //}

            // Not very useful as a general purpose method.
            static string SomeMethodThatMightThrow(string s) => s[4] == 'C' ? throw new InvalidOperationException() : $"""C:\newFolder\{s}""";

            string[] files = ["fileA.txt", "fileB.txt", "fileC.txt"];
            var exceptionDemoQuery = from file in files
                                     let n = SomeMethodThatMightThrow(file)
                                     select n;
            try
            {
                foreach (var item in exceptionDemoQuery)
                {
                    Console.WriteLine($"Processing {item}");
                }
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e.Message);
            }

            //var orderedQuery = from department in departments
            //                   join student in students on department.ID equals
            //                   student.DepartmentID into studentGroup
            //                   orderby department.Name
            //                   select new
            //                   {
            //                       DepartmentName = department.Name,
            //                       Students = from student in studentGroup
            //                                  orderby student.LastName
            //                                  select student
            //                   };

            List<string> phrases = ["an apple a day", "the quick brown fox"];
            var query = phrases.SelectMany(phrases => phrases.Split(' '));
            foreach (string s in query)
            {
                Console.WriteLine(s);
            }

            int[] array = [1, 2, 3, 4, 5, 6, 7, 8, 9];
            var queryMany = from i in array
                            from j in array
                            where j <= i
                            select $"{(j == 1 ? "\n" : "")}{j}*{i}={i * j}\t";
            Console.WriteLine(string.Concat(queryMany));

            var _queryMany = array
                                .SelectMany(i => array
                                    .Where(j => j <= i)
                                    .Select(j => $"{(j == 1 ? "\n" : "")}{j}*{i}={i * j}\t")
                                );

            Console.WriteLine(string.Concat(_queryMany));
        }
    }
}
