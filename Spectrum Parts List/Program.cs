using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Spectrum_Parts_List
{
    internal class Program
    {
        private const string FILEPATH = @"C:\Users\jon.wilson\source\repos\Spectrum Parts List\Spectrum Parts List\Files\PurchasingData.csv";
        private const string CSVFILEPATH = @"C:\Users\jon.wilson\source\repos\Spectrum Parts List\Spectrum Parts List\Files\FormattedPurchasingData.csv";
        private const string EXCEL1FILEPATH = @"C:\Users\jon.wilson\source\repos\Spectrum Parts List\Spectrum Parts List\Files\FormattedPurchasingData1.xlsx";
        private const string EXCEL2FILEPATH = @"C:\Users\jon.wilson\source\repos\Spectrum Parts List\Spectrum Parts List\Files\FormattedPurchasingData2.xlsx";
        private const string EXCEL3FILEPATH = @"C:\Users\jon.wilson\source\repos\Spectrum Parts List\Spectrum Parts List\Files\FormattedPurchasingData3.xlsx";

        static void Main(string[] args)
        {
            List<Part> parts = Utility.LoadPartsFromFile(FILEPATH)
                                .Distinct()
                                .OrderBy(p => p.item_code)
                                .ToList();
            Utility.SaveToExcel(parts, EXCEL1FILEPATH);

            //PartsWithCCandDescription – Group by PartNumber, CostCode, Description, Sum of Qty, Sum of TotalCost
            var partTotals_By_ItemCode_CC_Description = parts
                                            .GroupBy(p => new { p.item_code, p.CostCode, p.item_description })
                                            .Select(r => new {
                                                ItemCode = r.Key.item_code,
                                                Cost_Code = r.Key.CostCode,
                                                Description = r.Key.item_description,
                                                TotalQuantity = r.Sum(p => p.po_quantity_list1),
                                                TotalPrice = r.Sum(p => p.item_price)
                                            }).ToList();
            parts.Clear();
            partTotals_By_ItemCode_CC_Description.ForEach(r =>
                {
                    parts.Add(new Part(r.ItemCode, r.Cost_Code, r.Description, r.TotalQuantity, r.TotalPrice)); 
                }
            );
            Utility.SaveToExcel(parts, EXCEL2FILEPATH);

            //DistinctParts – Group by PartNumber only, Sum of Qty, Sum of TotalCost
            var partsTotals_By_ItemCode = parts
                                            .GroupBy(p => new { p.item_code })
                                            .Select(r => new {
                                                ItemCode = r.Key.item_code,
                                                TotalQuantity = r.Sum(p => p.po_quantity_list1),
                                                TotalPrice = r.Sum(p => p.item_price)
                                            }).ToList();
            parts.Clear();
            partsTotals_By_ItemCode.ForEach(r =>
                { 
                    parts.Add(new Part(r.ItemCode, r.TotalQuantity, r.TotalPrice)); 
                }
            );
            Utility.SaveToExcel(parts, EXCEL3FILEPATH);

            Console.WriteLine($"\nProcess complete. There are {parts.Count} distinct parts. Press any key to quit.");
            Console.ReadLine();
        }
    }
}
