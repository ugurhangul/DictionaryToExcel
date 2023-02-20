using Ougha.Entities;
using System.Collections.Generic;

namespace Ougha
{
    public static class DictionaryToExcelExtensions
    {
        public static byte[] ToExcel<T>(this T items, string sheetName = null) where T : List<Dictionary<string, string>>
        {
            return DictionaryToExcel.CreateExcel(items);
        }
        public static byte[] ToExcel<T>(this T items) where T : List<TabSheet>
        {
            return DictionaryToExcel.CreateExcel(items);
        }
    }
}