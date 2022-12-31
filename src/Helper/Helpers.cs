using SimpleExcelGenerator.Atributtes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace SimpleExcelGenerator.Helper
{
    internal static class Helpers
    {
        public static DataTable ToDataTable<T>(this IEnumerable<T> data, string name)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(data.First().GetType());
            DataTable table = new(name);

            foreach (PropertyDescriptor prop in properties)
            {
                ExcelColumnAttr memberAttr = prop.GetAttribute<ExcelColumnAttr>();
                table.Columns.Add(memberAttr != null ? memberAttr.DisplayName : prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties) {
                    ExcelColumnAttr memberAttr = prop.GetAttribute<ExcelColumnAttr>();
                    row[memberAttr != null ? memberAttr.DisplayName : prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                    
                table.Rows.Add(row);
            }
            return table;
        }

        public static T GetAttribute<T>(this PropertyDescriptor prop) where T : Attribute
        {
            foreach (Attribute att in prop.Attributes)
            {
                var tAtt = att as T;
                if (tAtt != null) return tAtt;
            }
            return null;
        }

        public static String BytesToString(this long byteCount)
        {
            string[] suf = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }
    }
}
