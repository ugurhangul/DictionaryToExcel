# Ougha.DictionaryToExcel [![NuGet version](https://badge.fury.io/nu/Ougha.DictionaryToExcel.svg)](https://badge.fury.io/nu/Ougha.DictionaryToExcel)

Create Excel from Dictionary
Forked from https://github.com/mustaddon/ArrayToExcel Thanks to Leonid Salavatov

### Example #1 for Umbraco Properties

```C# 
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

