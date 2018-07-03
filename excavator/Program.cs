using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excavator.coreEngine.Contracts;
using excavator.coreEngine.Impl;

namespace excavator {
    class Program {

        static void Main(string[] args) {
            IExcavator excavator = new Excavator();
            Console.WriteLine("Excavation starts....");
            excavator.FilePath = $@"{AppDomain.CurrentDomain.BaseDirectory}\data\Jan-Oct2012.xlsx";
            var date = new DateTime(2012, 1, 3);
            var data = excavator.GetStockDetailsByDate(date);
            Console.WriteLine($"{data.Count} rows found.");
            if (data != null && data.Any()) {
                var outputFile = $@"{AppDomain.CurrentDomain.BaseDirectory}\data\{date.ToString("yyyyMMdd")}\{date.ToString("yyyyMMdd")}.xlsx";

                excavator.WriteFile(data, outputFile);
            }
            Console.WriteLine($"{data.Count} rows write in new file.");
            Console.Read();
        }
    }
}
