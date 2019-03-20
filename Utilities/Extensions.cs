using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ClosedXML.Excel;

namespace Utilities
{
    public static class Extensions
    {
        public static void SaveToExcel<T>(this IEnumerable<T> collection, string fileName = "ExcelOutput", IEnumerable<string> removeColumns = default(IEnumerable<string>))
        {
            try
            {
                using (var excelWorkBook = collection.ExportToExcel(removeColumns))
                {
                    excelWorkBook.SaveAs($"{fileName}.xlsx");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static XLWorkbook ExportToExcel<T>(this IEnumerable<T> collection, IEnumerable<string> removeColumns = default(IEnumerable<string>))
        {
            var dataTable = collection.ExportToDataTable(removeColumns);
            using (var workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add(dataTable);
                workSheet.Rows().Style.Alignment.WrapText = false;
                return workBook;
            }
        }

        public static DataTable ExportToDataTable<T>(this IEnumerable<T> collection, IEnumerable<string> removeColumns = default(IEnumerable<string>))
        {
            var dataTable = new DataTable(typeof(T).Name);
            var properties = typeof(T).GetProperties();

            if (removeColumns != null)
            {
                properties = properties.Where(p => !removeColumns.Contains(p.Name)).ToArray();
            }
            foreach (var property in properties)
            {
                dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }
            foreach (var item in collection)
            {
                var rowData = properties.Select(p => p.GetValue(item, null) ?? DBNull.Value).ToArray();
                dataTable.Rows.Add(rowData);
            }
            return dataTable;
        }
    }
}
