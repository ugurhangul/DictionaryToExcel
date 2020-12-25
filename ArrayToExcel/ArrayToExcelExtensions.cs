using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace RandomSolutions
{
    public static class ArrayToExcelExtensions
    {
        public static byte[] ToExcel<T>(this T items, string sheetName = null) where T : Dictionary<string, string>
        {
            return ArrayToExcel.CreateExcel(items, scheme =>
            {
                scheme.SheetName = sheetName;
            });
        }

        public static byte[] ToExcel<T>(this T items, Action<ArrayToExcelScheme<T>> schemeBuilder)where T : Dictionary<string, string>
        {
            return ArrayToExcel.CreateExcel(items, schemeBuilder);
        }

    }
}
