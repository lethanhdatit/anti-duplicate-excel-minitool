using IronXL;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;

bool excelProcess = true;
bool jsonProcess = true;
bool autoSizeRow = true;
bool autoSizeColumn = true;
string path = "C:\\Users\\Admin\\Desktop\\duplicate\\";
string version = "v2";

string i_excelPath = $"{path}en_flat.xlsx";
string o_excelPath = $"{path}en_flat_unique_{version}.xlsx";
string t_excelPath = $"{path}__temp__.xlsx";
string i_jsonPath = $"{path}en_flat.json";
string o_jsonPath = $"{path}en_flat_unique_{version}.json";
string diff_jsonPath = $"{path}en_flat_diff_{version}.json";
string dup_jsonPath = $"{path}en_flat_dup_{version}.json";
string common_fe_prefix = "common";
string common_be_prefix = "common.backend";
string beKeyword = "backendService";
int levelPrefix = 4;

if (!excelProcess && !jsonProcess)
    return;

var wb_original = WorkBook.Load(i_excelPath);
if(wb_original != null)
{
    if (File.Exists(t_excelPath))
        File.Delete(t_excelPath);

    var wbTemp = wb_original.SaveAs(t_excelPath);
    wb_original.Close();

    File.Delete(wbTemp.FilePath);

    var wsTemp = wbTemp.GetWorkSheet("en") ?? wbTemp.DefaultWorkSheet;

    int colValueIndex = 2; // en   

    // keep current char case
    var uniqueListValues = wsTemp.Columns[colValueIndex]
                             .Where(s => s.RowIndex != 0) // not count header row
                             .GroupBy(g => g.Value);

    WorkBook? rWb = excelProcess ? WorkBook.Create(ExcelFileFormat.XLSX) : null;
    WorkSheet? rWs = null;
    if (rWb != null)
    {
        rWb.DefaultWorkSheet.Name = wsTemp.Name;
        rWs = rWb.DefaultWorkSheet;

        rWs["A1"].Value = "assign_to";
        rWs["A1"].Style.Font.Bold = true;

        rWs["B1"].Value = "new_key_name";
        rWs["B1"].Style.Font.Bold = true;

        rWs["C1"].Value = "duplicated_keys";
        rWs["C1"].Style.Font.Bold = true;

        rWs["D1"].Value = "en_unique";
        rWs["D1"].Style.Font.Bold = true;
    }
   
    int i = 2;
    List<Item> duplicatedFeKeyGroups = new List<Item>();
    List<Item> duplicatedBeKeyGroups = new List<Item>();
    foreach (var item in uniqueListValues)
    {
        List<string> duplicatedFeKeys = new List<string>();
        List<string> duplicatedBeKeys = new List<string>();
        foreach (var ii in item)
        {
            var key = wsTemp[$"B{ii.RowIndex + 1}"].StringValue;
           
            if (!string.IsNullOrWhiteSpace(key))
            {
                if(key.Contains(beKeyword, StringComparison.OrdinalIgnoreCase))
                    duplicatedBeKeys.Add(key);
                else
                    duplicatedFeKeys.Add(key);
            }
        }

        string newFeKey = string.Empty;
        if (duplicatedFeKeys.Any())
        {
            newFeKey = BuildNewKey(common_fe_prefix, duplicatedFeKeys, levelPrefix);

            if (jsonProcess)
                duplicatedFeKeyGroups.Add(new Item { NewKey = newFeKey, DuplicatedKeys = duplicatedFeKeys, Value = item.Key });
        }

        string newBeKey = string.Empty;
        if (duplicatedBeKeys.Any())
        {
            newBeKey = BuildNewKey(common_be_prefix, duplicatedBeKeys, levelPrefix);
            duplicatedBeKeyGroups.Add(new Item { NewKey = newBeKey, DuplicatedKeys = duplicatedBeKeys, Value = item.Key });
        }

        if (excelProcess)
        {
            rWs[$"B{i}"].Value = newFeKey;
            rWs[$"C{i}"].Value = string.Join(Environment.NewLine, duplicatedFeKeys);
            rWs[$"D{i}"].Value = item.Key;

            rWs[$"C{i}"].Style.WrapText = true;

            if (autoSizeRow)
                rWs.AutoSizeRow(i - 1);
        }

        i++;
    }

    if (excelProcess)
    {
        foreach (var be in duplicatedBeKeyGroups)
        {
            rWs[$"B{i}"].Value = be.NewKey;
            rWs[$"C{i}"].Value = string.Join(Environment.NewLine, be.DuplicatedKeys);
            rWs[$"D{i}"].Value = be.Value;

            rWs[$"C{i}"].Style.WrapText = true;

            if (autoSizeRow)
                rWs.AutoSizeRow(i - 1);

            i++;
        }

        if (autoSizeColumn)
        {
            rWs.AutoSizeColumn(0);
            rWs.AutoSizeColumn(1);
            rWs.AutoSizeColumn(2);
            rWs.AutoSizeColumn(3);
        }

        if (File.Exists(o_excelPath))
            File.Delete(o_excelPath);

        rWb?.SaveAs(o_excelPath);
        rWb?.Close();
    }

    wbTemp.Close();

    if (jsonProcess)
    {
        if (duplicatedFeKeyGroups.Any() || duplicatedBeKeyGroups.Any())
        {
            //var jObj = JsonConvert.DeserializeObject<JObject>(File.ReadAllText(i_jsonPath));
            long count = 0;
            JObject combines = new JObject();

            var dupkey = duplicatedFeKeyGroups.GroupBy(g => g.NewKey).Where(a => a.Count() > 1);
            var c = dupkey.Count();

            foreach (var gr in duplicatedFeKeyGroups)
                combines.Add(gr.NewKey, gr.Value.ToString());
            foreach (var gr in duplicatedBeKeyGroups)
                combines.Add(gr.NewKey, gr.Value.ToString());

            File.WriteAllText(o_jsonPath, JsonConvert.SerializeObject(combines, Formatting.Indented));
            var beDups = duplicatedBeKeyGroups.SelectMany(k => k.DuplicatedKeys);
            var feDups = duplicatedFeKeyGroups.SelectMany(s => s.DuplicatedKeys);
            var dups = feDups.Concat(beDups);
            File.WriteAllText(dup_jsonPath, JsonConvert.SerializeObject(new { total = dups.Count(), backend = beDups.Count(), frontend = feDups.Count(), result = dups }, Formatting.Indented));
        }
    }
}

static string BuildNewKey(string prefix, List<string> duplicatedKeys, int levelPrefix)
{
    string newKey = duplicatedKeys.FirstOrDefault(f => !f.Contains(".")) ?? duplicatedKeys.Select(s => new { count = s.Count(c => c == '.'), key = s }).OrderByDescending(s => s.count).FirstOrDefault().key;

    if (newKey.Contains("."))
    {
        var splited = newKey.Split('.');
        var suffix = string.Empty;
        for (int i = 1; i <= levelPrefix; i++)
        {
            if(splited.Length - i >= 0)
            {
                suffix = splited[splited.Length - i] + (!string.IsNullOrWhiteSpace(suffix) ? "." : string.Empty) + suffix;
            }                
        }
        newKey = $"{prefix}.{suffix}";
    }

    return newKey;
}