using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excavator.coreEngine.Contracts;
using excavator.coreEngine.Model;
using OfficeOpenXml;

namespace excavator.coreEngine.Impl {
    public class Excavator : IExcavator {

        /// <summary>
        /// Gets or sets the file path.
        /// </summary>
        /// <value>
        /// The file path.
        /// </value>
        public string FilePath { get; set; }

        /// <summary>
        /// Gets the stock details by date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <returns></returns>
        public IReadOnlyList<StockPriceModel> GetStockDetailsByDate(DateTime date) {

            using (var excel = new ExcelPackage(new FileInfo(FilePath))) {
                var worksheet = excel.Workbook.Worksheets.FirstOrDefault();
                var totalRows = worksheet?.Dimension.End.Row;
                var data = new List<StockPriceModel>();
                var dateRow = date.ToString("yyyyMMdd");

                for (var rowIndex = 1; rowIndex <= totalRows; rowIndex++) {
                    if (worksheet.Cells[rowIndex, 2].Text == dateRow) {
                        data.Add(new StockPriceModel() {
                            Name = worksheet.Cells[rowIndex, 1].Text,
                            TradingDate = worksheet.Cells[rowIndex, 2].Text,
                            TradingTime = worksheet.Cells[rowIndex, 3].Text,
                            OpenPrice = Convert.ToDouble(worksheet.Cells[rowIndex, 4].Text),
                            HighPrice = Convert.ToDouble(worksheet.Cells[rowIndex, 5].Text),
                            LowPrice = Convert.ToDouble(worksheet.Cells[rowIndex, 6].Text),
                            ClosePrice = Convert.ToDouble(worksheet.Cells[rowIndex, 7].Text)
                        });
                    }
                }
                return data;
            }
        }

        /// <summary>
        /// Writes the file.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="outputFile">The output file.</param>
        /// <returns></returns>
        public bool WriteFile(IReadOnlyList<StockPriceModel> data, string outputFile) {
            using (var excel = new ExcelPackage()) {

                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Sheet1");
                var rowIndex = 1;

                foreach (var row in data) {
                    worksheet.Cells[$"A{rowIndex}"].Value = row.Name;
                    worksheet.Cells[$"B{rowIndex}"].Value = row.TradingDate;
                    worksheet.Cells[$"C{rowIndex}"].Value = row.TradingTime;
                    worksheet.Cells[$"D{rowIndex}"].Value = row.OpenPrice;
                    worksheet.Cells[$"E{rowIndex}"].Value = row.HighPrice;
                    worksheet.Cells[$"F{rowIndex}"].Value = row.LowPrice;
                    worksheet.Cells[$"G{rowIndex}"].Value = row.ClosePrice;
                    rowIndex++;
                }

                var fileName = new FileInfo(outputFile);
                if (!fileName.Directory.Exists)
                {
                    Directory.CreateDirectory(fileName.Directory.FullName);
                }

                excel.SaveAs(fileName);
            }
            return false;
        }
    }
}
