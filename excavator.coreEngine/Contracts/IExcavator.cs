using System;
using System.Collections.Generic;
using excavator.coreEngine.Model;

namespace excavator.coreEngine.Contracts {
    public interface IExcavator {

        /// <summary>
        /// Gets or sets the file path.
        /// </summary>
        /// <value>
        /// The file path.
        /// </value>
        string FilePath { get; set; }

        /// <summary>
        /// Gets the stock details by date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <returns></returns>
        IReadOnlyList<StockPriceModel> GetStockDetailsByDate(DateTime date);

        /// <summary>
        /// Writes the file.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="outputFile">The output file.</param>
        /// <returns></returns>
        bool WriteFile(IReadOnlyList<StockPriceModel> data, string outputFile);
    }
}
