using System;
using System.Collections.Generic;

namespace Ougha
{
    public static class DictionaryToExcelExtensions
    {
        public static byte[] ToExcel<T>(this T items, string sheetName = null) where T : List<Dictionary<string, string>>
        {
            return DictionaryToExcel.CreateExcel(items, scheme =>
            {
                scheme.SheetName = sheetName;
            });
        }

        public static byte[] ToExcel<T>(this T items, Action<DictionaryToExcelScheme<T>> schemeBuilder) where T : List<Dictionary<string, string>>
        {
            return DictionaryToExcel.CreateExcel(items, schemeBuilder);
        }

    }
}
