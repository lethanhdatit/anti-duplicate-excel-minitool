using IronXL;
using System;

public class Item
{
    public string NewKey { get; set; }
    public List<string> DuplicatedKeys { get; set; }
    public object Value { get; set; }
}

public class StringIEqualityComparer : IEqualityComparer<string>
{
    public bool Equals(string x, string y)
    {
        return x.Equals(y, StringComparison.OrdinalIgnoreCase);
    }

    public int GetHashCode(string obj)
    {
        return obj.GetHashCode();
    }
}
