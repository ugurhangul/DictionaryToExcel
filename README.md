# ArrayToExcel [![NuGet version](https://badge.fury.io/nu/ArrayToExcel.svg)](http://badge.fury.io/nu/ArrayToExcel)
Create Excel from Array

### Example #1

```C# for Umbraco Properties
var list = new List<Dictionary<string, string>>();

foreach (var child in children.ToList())
{
    var dic = new Dictionary<string, string>();
    dic.Add("Name", child.Name);
    child.Properties.ForEach(c => dic.Add(c.PropertyType.Name, (c.GetValue() == null ? "-" : c.GetValue().ToString())));
    list.Add(dic);
}

var excel = list.ToExcel();
```

### Example #2

```C#
var list = new List<Dictionary<string, string>>();
var dictionary = new Dictionary<string, string>();

dictionary.Add("HeaderName", "Value");
list.Add(dictionary);

var excel = list.ToExcel();
```

