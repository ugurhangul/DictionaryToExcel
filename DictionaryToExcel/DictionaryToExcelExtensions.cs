using System;
using System.Collections.Generic;

namespace Ougha
{
    public static class DictionaryToExcelExtensions
    {
        public static byte[] ToExcel<T>(this T items, string sheetName = null) where T : Dictionary<string, string>
        {
            return DictionaryToExcelToExcel.CreateExcel(items, scheme =>
            {
                scheme.SheetName = sheetName;
            });
        }

        public static byte[] ToExcel<T>(this T items, Action<DictionaryToExcelScheme<T>> schemeBuilder)where T : Dictionary<string, string>
        {
            return DictionaryToExcelToExcel.CreateExcel(items, schemeBuilder);
        }

    }
}
