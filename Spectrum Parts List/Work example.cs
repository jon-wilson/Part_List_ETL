using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Spectrum_Parts_List
{
    internal class Program
    {
        private const string FILEPATH = @"~\Files\PurchasingData.csv";
        private const string CSVFILEPATH = @"~\Files\FormattedPurchasingData.csv";
        private const string EXCEL1FILEPATH = @"~\Files\FormattedPurchasingData1.xlsx";
        private const string EXCEL2FILEPATH = @"~\Files\FormattedPurchasingData2.xlsx";
        private const string EXCEL3FILEPATH = @"~\Files\FormattedPurchasingData3.xlsx";

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
	
	internal static class Utility
    {
        private const string FORMAT = "Date: {0} | Method: {1} | Row #: {2} | Message {3}";
        private const string LOGFILE = "Log.txt";

        public static void LogWriter(string method, string rowNum, string message)
        {
            IEnumerable<string> text = new List<string>() {
                string.Format(FORMAT, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), method, rowNum, message)
            };
            File.AppendAllLines(LOGFILE, text);
        }

        public static List<Part> LoadPartsFromFile(string filepath)
        {
            List<Part> parts = new List<Part>();
            int i = 1;

            using (var reader = new StreamReader(filepath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (i > 1)
                    {                        
                        var values = line.Split(',');

                        if (values.Length == 28)
                            parts.Add(new Part(values, i));
                        else
                            Utility.LogWriter("LoadPartsFromFile()", i.ToString(), "Incorrect number of columns due to extra commas in record");
                    }
                    i++;
                }
            }

            return parts;
        }

        public static void SaveToCSV(List<Part> parts, string csvFilePath)
        {
            var lines = new List<string>();
            IEnumerable<PropertyDescriptor> props = TypeDescriptor.GetProperties(typeof(Part)).OfType<PropertyDescriptor>();
            var header = string.Join(",", props.ToList().Select(x => x.Name));
            lines.Add(header);
            var valueLines = parts.Select(row => string.Join(",", header.Split(',').Select(a => row.GetType().GetProperty(a).GetValue(row, null))));
            lines.AddRange(valueLines);
            File.WriteAllLines(csvFilePath, lines.ToArray());
        }

        public static void SaveToExcel(List<Part> parts, string excelFilePath)
        {
            var workbook = new XLWorkbook();
            DataTable dtParts = new DataTable(typeof(Part).Name);

            //Get all the properties by using reflection   
            PropertyInfo[] Props = typeof(Part).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                dtParts.Columns.Add(prop.Name);
            }

            foreach (Part part in parts)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(part, null);
                }
                dtParts.Rows.Add(values);
            }

            workbook.AddWorksheet(dtParts);
            workbook.SaveAs(excelFilePath);
        }
    }	
	
    internal class Part : IEquatable<Part>
    {
        #region Properties
        public string company_code { get; set; }
        public string po_number { get; set; }
        public string PONumber { get; set; }
        public DateTime PODate { get; set; }
        public int line_number { get; set; }
        public decimal po_quantity_list1 { get; set; }
        public decimal po_quantity_list2 { get; set; }
        public string item_code { get; set; }
        public string PartNumber { get; set; }
        public string item_description { get; set; }
        public string unit_of_measure { get; set; } 
        public decimal item_price { get; set; }
        public decimal line_extension_list1 { get; set; }
        public decimal line_extension_list2 { get; set; }
        public DateTime delivery_date { get; set; }
        public string gl_account { get; set; }
        public string job_number { get; set; }
        public string phase_code { get; set; }
        public string cost_type { get; set; }
        public decimal received_extension { get; set; }
        public decimal OpenAmount { get; set; }
        public string Job { get; set; }
        public string JobName { get; set; }
        public string vendor_code { get; set; }
        public string VendorName { get; set; }
        public string CostCode { get; set; }
        public string AECostCode { get; set; }
        public string AECostCodeCategory { get; set; }
        #endregion

        #region Constructors
        public Part(string[] part, int rowNum)
        {
            try
            {
                this.company_code = part[0].Trim();
                this.po_number = part[1].Trim();
                PONumber = part[2].Trim();
                PODate = Convert.ToDateTime(part[3].Trim());
                this.line_number = Convert.ToInt32(part[4].Trim());
                this.po_quantity_list1 = Convert.ToDecimal(part[5].Trim());
                this.po_quantity_list2 = Convert.ToDecimal(part[6].Trim());
                this.item_code = Regex.Replace(part[7].Trim(), @"^[^a-zA-Z0-9]+", String.Empty); 
                PartNumber = Regex.Replace(part[8].Trim(), @"^[^a-zA-Z0-9]+", String.Empty);
                this.item_description = part[9].Trim();
                this.unit_of_measure = part[10].Trim();
                this.item_price = Convert.ToDecimal(part[11].Trim());
                this.line_extension_list1 = Convert.ToDecimal(part[12].Trim());
                this.line_extension_list2 = Convert.ToDecimal(part[13].Trim());
                this.delivery_date = part[14].Equals("") ? DateTime.MinValue : Convert.ToDateTime(part[14].Trim());
                this.gl_account = part[15].Trim();
                this.job_number = part[16].Trim();
                this.phase_code = part[17].Trim();
                this.cost_type = part[18].Trim();
                this.received_extension = Convert.ToDecimal(part[19].Trim());
                OpenAmount = Convert.ToDecimal(part[20].Trim());
                Job = part[21].Trim();
                JobName = part[22].Trim();
                this.vendor_code = part[23].Trim();
                VendorName = part[24].Trim();
                CostCode = part[25].Trim();
                AECostCode = part[26].Trim();
                AECostCodeCategory = part[27].Trim();
            }
            catch (Exception e)
            {
                Utility.LogWriter("Part constructor", rowNum.ToString(), e.Message);
            }
        }

        public Part(string item_code, string costCode, string item_description, decimal po_quantity_list1, decimal item_price)
        {
            this.item_code = item_code;
            CostCode = costCode;
            this.item_description = item_description;
            this.po_quantity_list1 = po_quantity_list1;
            this.item_price = item_price;
        }

        public Part(string item_code, decimal po_quantity_list1, decimal item_price)
        {
            this.item_code = item_code;
            this.po_quantity_list1 = po_quantity_list1;
            this.item_price = item_price;
        }

        public Part(string company_code, string po_number, string pONumber, DateTime pODate, int line_number, int po_quantity_list1, int po_quantity_list2, string item_code, string partNumber, string item_description, string unit_of_measure, decimal item_price, decimal line_extension_list1, decimal line_extension_list2, DateTime delivery_date, string gl_account, string job_number, string phase_code, string cost_type, decimal received_extension, decimal openAmount, string job, string jobName, string vendor_code, string vendorName, string costCode, string aECostCode, string aECostCodeCategory)
        {
            this.company_code = company_code;
            this.po_number = po_number;
            PONumber = pONumber;
            PODate = pODate;
            this.line_number = line_number;
            this.po_quantity_list1 = po_quantity_list1;
            this.po_quantity_list2 = po_quantity_list2;
            this.item_code = Regex.Replace(item_code, @"^[^a-zA-Z0-9]+", String.Empty);
            PartNumber = Regex.Replace(partNumber, @"^[^a-zA-Z0-9]+", String.Empty);
            this.item_description = item_description;
            this.unit_of_measure = unit_of_measure;
            this.item_price = item_price;
            this.line_extension_list1 = line_extension_list1;
            this.line_extension_list2 = line_extension_list2;
            this.delivery_date = delivery_date;
            this.gl_account = gl_account;
            this.job_number = job_number;
            this.phase_code = phase_code;
            this.cost_type = cost_type;
            this.received_extension = received_extension;
            OpenAmount = openAmount;
            Job = job;
            JobName = jobName;
            this.vendor_code = vendor_code;
            VendorName = vendorName;
            CostCode = costCode;
            AECostCode = aECostCode;
            AECostCodeCategory = aECostCodeCategory;
        }
        #endregion

        #region Methods
        public bool Equals(Part other)
        {
            if (object.ReferenceEquals(other, null))
            {
                return false;
            }
            if (object.ReferenceEquals(this, other))
            {
                return true;
            }
            return this.company_code.Equals(other.company_code) &&
                this.po_number.Equals(other.po_number) &&
                PONumber.Equals(other.PONumber) &&
                PODate.Equals(other.PODate) &&
                this.line_number.Equals(other.line_number) &&
                this.po_quantity_list1.Equals(other.po_quantity_list1) &&
                this.po_quantity_list2.Equals(other.po_quantity_list2) &&
                this.item_code.Equals(other.item_code) &&
                PartNumber.Equals(other.PartNumber) &&
                this.item_description.Equals(other.item_description) &&
                this.unit_of_measure.Equals(other.unit_of_measure) &&
                this.item_price.Equals(other.item_price) &&
                this.line_extension_list1.Equals(other.line_extension_list1) &&
                this.line_extension_list2.Equals(other.line_extension_list2) &&
                this.delivery_date.Equals(other.delivery_date) &&
                this.gl_account.Equals(other.gl_account) &&
                this.job_number.Equals(other.job_number) &&
                this.phase_code.Equals(other.phase_code) &&
                this.cost_type.Equals(other.cost_type) &&
                this.received_extension.Equals(other.received_extension) &&
                OpenAmount.Equals(other.OpenAmount) &&
                Job.Equals(other.Job) &&
                JobName.Equals(other.JobName) &&
                this.vendor_code.Equals(other.vendor_code) &&
                VendorName.Equals(other.VendorName) &&
                CostCode.Equals(other.CostCode) &&
                AECostCode.Equals(other.AECostCode) &&
                AECostCodeCategory.Equals(other.AECostCodeCategory);
        }

        public override int GetHashCode()
        {             
            int hc_company_code = this.company_code.GetHashCode();
            int hc_po_number = this.po_number.GetHashCode();
            int hc_ponumber = PONumber.GetHashCode();
            int hc_podate = PODate.GetHashCode();
            int hc_line_number = this.line_number.GetHashCode();
            int hc_po_quantity_list1 = this.po_quantity_list1.GetHashCode();
            int hc_po_quantity_list2 = this.po_quantity_list2.GetHashCode();
            int hc_item_code = this.item_code.GetHashCode();
            int hc_partnumber = PartNumber.GetHashCode();
            int hc_item_description = this.item_description.GetHashCode();
            int hc_unit_of_measure = this.unit_of_measure.GetHashCode();
            int hc_item_price = this.item_price.GetHashCode();
            int hc_line_extension_list1 = this.line_extension_list1.GetHashCode();
            int hc_line_extension_list2 = this.line_extension_list2.GetHashCode();
            int hc_delivery_date = this.delivery_date.GetHashCode();
            int hc_gl_account = this.gl_account.GetHashCode();
            int hc_job_number = this.job_number.GetHashCode();
            int hc_phase_code = this.phase_code.GetHashCode();
            int hc_cost_type = this.cost_type.GetHashCode();
            int hc_received_extension = this.received_extension.GetHashCode();
            int hc_openamount = OpenAmount.GetHashCode();
            int hc_job = Job.GetHashCode();
            int hc_jobname = JobName.GetHashCode();
            int hc_vendor_code = this.vendor_code.GetHashCode();
            int hc_vendorname = VendorName.GetHashCode();
            int hc_costcode = CostCode.GetHashCode();
            int hc_aecostcode = AECostCode.GetHashCode();
            int hc_aecostcodecategory = AECostCodeCategory.GetHashCode();

            return hc_company_code ^
             hc_po_number ^ 
             hc_ponumber ^ 
             hc_podate ^
             hc_line_number ^
             hc_po_quantity_list1 ^ 
             hc_po_quantity_list2 ^ 
             hc_item_code ^
             hc_partnumber ^ 
             hc_item_description ^
             hc_unit_of_measure ^ 
             hc_item_price ^
             hc_line_extension_list1 ^
             hc_line_extension_list2 ^
             hc_delivery_date ^
             hc_gl_account ^
             hc_job_number ^
             hc_phase_code ^
             hc_cost_type ^
             hc_received_extension ^
             hc_openamount ^
             hc_job ^
             hc_jobname ^
             hc_vendor_code ^
             hc_vendorname ^
             hc_costcode ^
             hc_aecostcode ^
             hc_aecostcodecategory;
        }
        #endregion
    }
}
