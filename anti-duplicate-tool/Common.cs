﻿using IronXL;
using System;

public enum ItemStatus 
{
    NO_CHANGED,
    ADDED_NEW,
    CHANGED,
    REMOVED
}
public class Item
{
    public string Key { get; set; }
    public List<string> DuplicatedKeys { get; set; }
    public object Value { get; set; }
    public string CompareNote { get; set; }
    public ItemStatus CompareStatus { get; set; }
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

public static class Common
{
    public static string BuildNewKey(string prefix, List<string> duplicatedKeys, int levelPrefix)
    {
        string newKey = duplicatedKeys.FirstOrDefault(f => !f.Contains(".")) ?? duplicatedKeys.Select(s => new { count = s.Count(c => c == '.'), key = s }).OrderByDescending(s => s.count).FirstOrDefault().key;

        if (newKey.Contains("."))
        {
            var splited = newKey.Split('.');
            var suffix = string.Empty;
            for (int i = 1; i <= levelPrefix; i++)
            {
                if (splited.Length - i >= 0)
                {
                    suffix = splited[splited.Length - i] + (!string.IsNullOrWhiteSpace(suffix) ? "." : string.Empty) + suffix;
                }
            }
            newKey = $"{prefix}.{suffix}";
        }

        return newKey;
    }

    public static void ProcessMetadata(bool autoDoublePrefixForDupKey, string addressInCol, int levelPrefix, string common_fe_prefix, string common_be_prefix, string beKeyword, WorkSheet wsTemp, IEnumerable<IGrouping<object, Cell>> uniqueListValues, List<Item> duplicatedFeKeyGroups, List<Item> duplicatedBeKeyGroups, out List<IGrouping<string, Item>> dupkeyFe, out List<IGrouping<string, Item>> dupkeyBe)
    {
        foreach (var item in uniqueListValues)
        {
            List<string> duplicatedFeKeys = new List<string>();
            List<string> duplicatedBeKeys = new List<string>();
            foreach (var ii in item)
            {
                var key = wsTemp[$"{addressInCol}{ii.RowIndex + 1}"].StringValue;

                if (!string.IsNullOrWhiteSpace(key))
                {
                    if (key.Contains(beKeyword, StringComparison.OrdinalIgnoreCase))
                        duplicatedBeKeys.Add(key);
                    else
                        duplicatedFeKeys.Add(key);
                }
            }

            string newFeKey = string.Empty;
            if (duplicatedFeKeys.Any())
            {
                newFeKey = duplicatedFeKeys.Count > 1 ? Common.BuildNewKey(common_fe_prefix, duplicatedFeKeys, levelPrefix) : duplicatedFeKeys.FirstOrDefault();
                duplicatedFeKeyGroups.Add(new Item { Key = newFeKey, DuplicatedKeys = duplicatedFeKeys, Value = item.Key });
            }

            string newBeKey = string.Empty;
            if (duplicatedBeKeys.Any())
            {
                newBeKey = duplicatedBeKeys.Count > 1 ? Common.BuildNewKey(common_be_prefix, duplicatedBeKeys, levelPrefix) : duplicatedBeKeys.FirstOrDefault();
                duplicatedBeKeyGroups.Add(new Item { Key = newBeKey, DuplicatedKeys = duplicatedBeKeys, Value = item.Key });
            }
        }

        dupkeyFe = duplicatedFeKeyGroups.GroupBy(g => g.Key).Where(a => a.Count() > 1).ToList();

        foreach (var group in dupkeyFe)
        {
            duplicatedFeKeyGroups.RemoveAll(s => s.Key == group.Key);
            if (autoDoublePrefixForDupKey)
            {
                foreach (var item in group)
                {
                    var temp = item.DuplicatedKeys.Select(d =>
                    {
                        if (d.StartsWith(common_fe_prefix))
                            d = $"{common_fe_prefix}.{d}";
                        return d;
                    }).ToList();
                    item.DuplicatedKeys = temp;
                }
            }           
        }

        dupkeyBe = duplicatedBeKeyGroups.GroupBy(g => g.Key).Where(a => a.Count() > 1).ToList();
        foreach (var group in dupkeyBe)
        {
            duplicatedBeKeyGroups.RemoveAll(s => s.Key == group.Key);
            if (autoDoublePrefixForDupKey)
            {
                foreach (var item in group)
                {
                    var temp = item.DuplicatedKeys.Select(d =>
                    {
                        if (d.StartsWith(common_be_prefix))
                            d = $"{common_be_prefix}.{d}";
                        return d;
                    }).ToList();
                    item.DuplicatedKeys = temp;
                }
            }
        }
    }


    public static void ProcessCompare(WorkSheet? rWs, List<Item> currentKeyGroups, List<Item> compareResults, List<Item> compareWithKeyGroups)
    {
        List<Item> readOnlyItems = new List<Item>();

        foreach (var currentRow in currentKeyGroups)
        {
            var exitedValue = compareWithKeyGroups.FirstOrDefault(a => a.Value.ToString() == currentRow.Value.ToString());
            var exitedKey = compareWithKeyGroups.FirstOrDefault(a => a.Key == currentRow.Key);

            if (exitedValue == null)
            {
                currentRow.CompareNote += $"New VALUE detected at column '{rWs["E1"].Value}'" + Environment.NewLine;
                currentRow.CompareStatus = ItemStatus.ADDED_NEW;
            }

            if (exitedKey == null)
            {
                currentRow.CompareNote += $"New KEY detected at column '{rWs["C1"].Value}'" + Environment.NewLine;
                currentRow.CompareStatus = ItemStatus.ADDED_NEW;
            }

            if (exitedKey != null)
            {
                foreach (var cr in currentRow.DuplicatedKeys)
                {
                    var newDup = exitedKey.DuplicatedKeys.FirstOrDefault(f => f == cr);
                    if (newDup == null)
                    {
                        currentRow.CompareNote += $"'{cr}' added into column '{rWs["D1"].Value}'" + Environment.NewLine;
                        currentRow.CompareStatus = ItemStatus.CHANGED;
                    }
                }
                foreach (var old in exitedKey.DuplicatedKeys)
                {
                    var oldDup = currentRow.DuplicatedKeys.FirstOrDefault(f => f == old);
                    if (oldDup == null)
                    {
                        currentRow.CompareNote += $"'{old}' removed from column '{rWs["D1"].Value}'" + Environment.NewLine;
                        currentRow.CompareStatus = ItemStatus.CHANGED;
                    }
                }
            }
        }

        foreach (var compareWithRow in compareWithKeyGroups)
        {
            var exitedValue = currentKeyGroups.FirstOrDefault(a => a.Value.ToString() == compareWithRow.Value.ToString());
            var exitedKey = currentKeyGroups.FirstOrDefault(a => a.Key == compareWithRow.Key);

            Item temp = new Item
            {
                Key = compareWithRow.Key,
                DuplicatedKeys = compareWithRow.DuplicatedKeys,
                Value = compareWithRow.Value
            };

            if (exitedValue == null)
            {
                temp.CompareNote += $"VALUE not found in current version (column '{rWs["E1"].Value}')" + Environment.NewLine;
                temp.CompareStatus = ItemStatus.REMOVED;
            }

            if (exitedKey == null)
            {
                temp.CompareNote += $"KEY not found in current version (column '{rWs["C1"].Value}')" + Environment.NewLine;
                temp.CompareStatus = ItemStatus.REMOVED;
            }

            if (!string.IsNullOrWhiteSpace(temp.CompareNote))
                currentKeyGroups.Add(temp);
        }

        compareResults.AddRange(currentKeyGroups.Where(w => w.CompareStatus != ItemStatus.NO_CHANGED));
        currentKeyGroups.RemoveAll(w => w.CompareStatus != ItemStatus.NO_CHANGED);
    }
    
    public static void ProcessCompare(WorkSheet? rWs, List<IGrouping<string, Item>> currentKeyGroups, List<Item> compareResults, List<IGrouping<string, Item>> compareKeyGroups)
    {
        List<Item> readOnlyItems = new List<Item>();

        var compareKeyFlatGroups = compareKeyGroups.SelectMany(s => s).ToList();
        var currentKeyFlatGroups = currentKeyGroups.SelectMany(s => s).ToList();

        foreach (var currentRow in currentKeyFlatGroups)
        {
            var exitedValue = compareKeyFlatGroups.FirstOrDefault(a => a.Value.ToString() == currentRow.Value.ToString());
            var exitedKey = compareKeyFlatGroups.FirstOrDefault(a => a.Key == currentRow.Key);

            if (exitedValue == null)
            {
                currentRow.CompareNote += $"New VALUE detected at column '{rWs["E1"].Value}'" + Environment.NewLine;
                currentRow.CompareStatus = ItemStatus.ADDED_NEW;
            }

            if (exitedKey == null)
            {
                currentRow.CompareNote += $"New KEY detected at column '{rWs["C1"].Value}'" + Environment.NewLine;
                currentRow.CompareStatus = ItemStatus.ADDED_NEW;
            }
        }

        foreach (var compareWithRow in compareKeyFlatGroups)
        {
            var exitedValue = currentKeyFlatGroups.FirstOrDefault(a => a.Value.ToString() == compareWithRow.Value.ToString());
            var exitedKey = currentKeyFlatGroups.FirstOrDefault(a => a.Key == compareWithRow.Key);

            Item temp = new Item
            {
                Key = compareWithRow.Key,
                DuplicatedKeys = compareWithRow.DuplicatedKeys,
                Value = compareWithRow.Value
            };

            if (exitedValue == null)
            {
                temp.CompareNote += $"VALUE not found in current version (column '{rWs["E1"].Value}')" + Environment.NewLine;
                temp.CompareStatus = ItemStatus.REMOVED;
            }

            if (exitedKey == null)
            {
                temp.CompareNote += $"KEY not found in current version (column '{rWs["C1"].Value}')" + Environment.NewLine;
                temp.CompareStatus = ItemStatus.REMOVED;
            }

            if (!string.IsNullOrWhiteSpace(temp.CompareNote))
                currentKeyFlatGroups.Add(temp);
        }

        compareResults.AddRange(currentKeyFlatGroups.Where(w => w.CompareStatus != ItemStatus.NO_CHANGED));
        currentKeyFlatGroups.RemoveAll(w => w.CompareStatus != ItemStatus.NO_CHANGED);

        currentKeyGroups = currentKeyFlatGroups.GroupBy(w => w.Key).ToList();
    }

    public static string ColumnIndexToColumnLetter(int colIndex)
    {
        int div = colIndex;
        string colLetter = String.Empty;
        int mod = 0;

        while (div > 0)
        {
            mod = (div - 1) % 26;
            colLetter = (char)(65 + mod) + colLetter;
            div = (int)((div - mod) / 26);
        }
        return colLetter;
    }

    public static int ColumnLetterToColumnIndex(string columnLetter)
    {
        columnLetter = columnLetter.ToUpper();
        int sum = 0;

        for (int i = 0; i < columnLetter.Length; i++)
        {
            sum *= 26;
            sum += (columnLetter[i] - 'A' + 1);
        }
        return sum;
    }
}
