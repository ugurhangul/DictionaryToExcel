using System.Collections.Generic;

namespace Ougha
{
    public class DictionaryToExcelScheme<T>
    {
        //internal ArrayToExcelScheme()
        //{
        //    _initDefaultColumns();
        //}

        public string SheetName;

        public DictionaryToExcelScheme<T> AddColumn(string name, string value, uint width = _defaultWidth)
        {
            (_columns ?? (_columns = new List<Column>())).Add(new Column
            {
                Index = Columns.Count,
                Name = name,
                ValueFn = value,
                Width = width,
            });
            return this;
        }


        //void _initDefaultColumns()
        //{
        //    var members = typeof(T).GetMembers(BindingFlags.Instance | BindingFlags.Public)
        //         .Where(x => x is PropertyInfo || x is FieldInfo);

        //    foreach (var member in members)
        //        _defaultColumns.Add(new Column
        //        {
        //            Index = _defaultColumns.Count,
        //            Name = member.Name,
        //            ValueFn = (x => (member as PropertyInfo)?.GetValue(x) ?? (member as FieldInfo)?.GetValue(x).ToString()),
        //            Width = _defaultWidth,
        //        });
        //}

        const uint _defaultWidth = 20;
        List<Column> _defaultColumns = new List<Column>();
        List<Column> _columns;

        internal List<Column> Columns => _columns ?? _defaultColumns;

        internal class Column
        {
            public int Index { get; set; }
            public string Name { get; set; }
            public string ValueFn { get; set; }
            public uint Width { get; set; }
        }


    }
}
