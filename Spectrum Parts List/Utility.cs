using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Spectrum_Parts_List
{
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
}
