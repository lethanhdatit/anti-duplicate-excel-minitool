using IronXL;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;

bool excelProcess = true;
bool jsonProcess = true;
bool autoSizeRow = true;
bool autoSizeColumn = true;
int levelPrefix = 1;

string version = "v4";
string compareWithVersion = "v3";

string path = "C:\\Users\\Admin\\Desktop\\duplicate\\";
string prefixFileName = "en_flat";

string i_excelPath = $"{path}{prefixFileName}_{version}.xlsx";
string i_jsonPath = $"{path}{prefixFileName}_{version}.json";

string o_excelPath = $"{path}{prefixFileName}_unique_{version}.xlsx";
string o_jsonPath = $"{path}{prefixFileName}_unique_{version}.json";

string c_excelPath = $"{path}{prefixFileName}_{compareWithVersion}.xlsx";

string dup_jsonPath = $"{path}{prefixFileName}_dup_{version}.json";
string t_excelPath = $"{path}__temp__.xlsx";

string common_fe_prefix = "common";
string common_be_prefix = "common.backend";
string beKeyword = "backendService";

if (!excelProcess && !jsonProcess)
    return;

bool processCompare = !string.IsNullOrWhiteSpace(compareWithVersion);

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
    string checkDupTargetColumn = wsTemp.Columns[colValueIndex].Rows[0].Value.ToString();

    // keep current char case
    var uniqueListValues = wsTemp.Columns[colValueIndex]
                             .Where(s => s.RowIndex != 0) // not count header row
                             .GroupBy(g => g.Value.ToString());

    WorkBook? rWb = excelProcess ? WorkBook.Create(ExcelFileFormat.XLSX) : null;
    WorkSheet? rWs = null;
    if (rWb != null)
    {
        rWb.DefaultWorkSheet.Name = wsTemp.Name;
        rWs = rWb.DefaultWorkSheet;

        rWs["A1"].Value = "compare_status" + (processCompare ? $"_with_{compareWithVersion}" : string.Empty);
        rWs["A1"].Style.Font.Bold = true;

        rWs["B1"].Value = "compare_note" + (processCompare ? $"_with_{compareWithVersion}" : string.Empty);
        rWs["B1"].Style.Font.Bold = true;

        rWs["C1"].Value = "new/existed_key";
        rWs["C1"].Style.Font.Bold = true;

        rWs["D1"].Value = "duplicated_keys";
        rWs["D1"].Style.Font.Bold = true;

        rWs["E1"].Value = $"{checkDupTargetColumn}_unique";
        rWs["E1"].Style.Font.Bold = true;
    }

    int i = 2;
    List<Item> duplicatedFeKeyGroups = new List<Item>();
    List<Item> duplicatedBeKeyGroups = new List<Item>();
    List<Item> compareResults = new List<Item>();
    List<IGrouping<string, Item>> dupkeyFe, dupkeyBe;

    Common.ProcessMetadata(levelPrefix, common_fe_prefix, common_be_prefix, beKeyword, wsTemp, uniqueListValues, duplicatedFeKeyGroups, duplicatedBeKeyGroups, out dupkeyFe, out dupkeyBe);

    if (excelProcess)
    {
        if (processCompare)
        {
            if (!File.Exists(c_excelPath))
                Console.WriteLine($"Target to compare not found. path: {c_excelPath}");
            else
            {
                var cWb = WorkBook.Load(c_excelPath);
                var cWs = cWb.GetWorkSheet("en") ?? cWb.DefaultWorkSheet;

                List<Item> compareFeKeyGroups = new List<Item>();
                List<Item> compareBeKeyGroups = new List<Item>();
                List<IGrouping<string, Item>> compareDupkeyFe, compareDupkeyBe;

                var compareUniqueListValues = cWs.Columns[colValueIndex]
                                                 .Where(s => s.RowIndex != 0) // not count header row
                                                 .GroupBy(g => g.Value.ToString());

                Common.ProcessMetadata(levelPrefix, common_fe_prefix, common_be_prefix, beKeyword, cWs, compareUniqueListValues, compareFeKeyGroups, compareBeKeyGroups, out compareDupkeyFe, out compareDupkeyBe);

                Common.ProcessCompare(rWs, duplicatedFeKeyGroups, compareResults, compareFeKeyGroups);

                Common.ProcessCompare(rWs, duplicatedBeKeyGroups, compareResults, compareBeKeyGroups);
            }
        }

        rWs.AutoSizeRow(0);

        foreach (var cp in compareResults)
        {
            rWs[$"B{i}"].Value = cp.CompareNote;
            rWs[$"A{i}"].Value = cp.CompareStatus.ToString();
            rWs[$"C{i}"].Value = cp.Key;
            rWs[$"D{i}"].Value = string.Join(Environment.NewLine, cp.DuplicatedKeys);
            rWs[$"E{i}"].Value = cp.Value;

            rWs[$"B{i}"].Style.WrapText = true;
            rWs[$"D{i}"].Style.WrapText = true;

            if(cp.CompareStatus == ItemStatus.ADDED_NEW)
            {
                rWs[$"A{i}"].Style.Font.Color = "#05F50D";
                rWs[$"C{i}:E{i}"].Style.Font.Color = "#05F50D";
            }
            else if(cp.CompareStatus == ItemStatus.REMOVED)
            {
                rWs[$"A{i}"].Style.Font.Color = "#F50505";
                rWs[$"C{i}:E{i}"].Style.Font.Color = "#F50505";
            }
            else if (cp.CompareStatus == ItemStatus.CHANGED)
            {
                rWs[$"A{i}"].Style.Font.Color = "#F57D05";
                rWs[$"D{i}"].Style.Font.Color = "#F57D05";
            }

            if (autoSizeRow)
                rWs.AutoSizeRow(i - 1);

            i++;
        }

        foreach (var fe in duplicatedFeKeyGroups)
        {
            rWs[$"C{i}"].Value = fe.Key;
            rWs[$"D{i}"].Value = string.Join(Environment.NewLine, fe.DuplicatedKeys);
            rWs[$"E{i}"].Value = fe.Value;

            rWs[$"D{i}"].Style.WrapText = true;

            if (autoSizeRow)
                rWs.AutoSizeRow(i - 1);

            i++;
        }

        foreach (var gr in dupkeyFe)
            foreach (var item in gr)
                foreach (var d in item.DuplicatedKeys)
                {
                    rWs[$"C{i}"].Value = d;
                    rWs[$"D{i}"].Value = "skipped process because duplicated new key -> accept duplicate value";
                    rWs[$"E{i}"].Value = item.Value.ToString();

                    rWs[$"D{i}"].Style.WrapText = true;

                    if (autoSizeRow)
                        rWs.AutoSizeRow(i - 1);

                    i++;
                }

        foreach (var gr in dupkeyBe)
            foreach (var item in gr)
                foreach (var d in item.DuplicatedKeys)
                {
                    rWs[$"C{i}"].Value = d;
                    rWs[$"D{i}"].Value = "skipped process because duplicated new key -> accept duplicate value";
                    rWs[$"E{i}"].Value = item.Value.ToString();

                    rWs[$"D{i}"].Style.WrapText = true;

                    if (autoSizeRow)
                        rWs.AutoSizeRow(i - 1);

                    i++;
                }

        foreach (var be in duplicatedBeKeyGroups)
        {
            rWs[$"C{i}"].Value = be.Key;
            rWs[$"D{i}"].Value = string.Join(Environment.NewLine, be.DuplicatedKeys);
            rWs[$"E{i}"].Value = be.Value;

            rWs[$"D{i}"].Style.WrapText = true;

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
        if (duplicatedFeKeyGroups.Any() || duplicatedBeKeyGroups.Any() || compareResults.Any())
        {
            JObject combines = new JObject();

            foreach (var gr in compareResults.Where(w => w.CompareStatus != ItemStatus.REMOVED))
                combines.Add(gr.Key, gr.Value.ToString());

            foreach (var gr in duplicatedFeKeyGroups)
                combines.Add(gr.Key, gr.Value.ToString());

            foreach (var group in dupkeyFe)
                foreach (var item in group)
                    foreach (var d in item.DuplicatedKeys)
                    {
                        if (!combines.ContainsKey(d))
                        {
                            combines.Add(d, item.Value.ToString());
                        }
                    }

            foreach (var group in dupkeyBe)
                foreach (var item in group)
                    foreach (var d in item.DuplicatedKeys)
                    {
                        if (!combines.ContainsKey(d))
                        {
                            combines.Add(d, item.Value.ToString());
                        }
                    }

            foreach (var gr in duplicatedBeKeyGroups)
                combines.Add(gr.Key, gr.Value.ToString());

            File.WriteAllText(o_jsonPath, JsonConvert.SerializeObject(combines, Formatting.Indented));
            var beDups = duplicatedBeKeyGroups.SelectMany(k => k.DuplicatedKeys);
            var feDups = duplicatedFeKeyGroups.SelectMany(s => s.DuplicatedKeys);
            var dups = feDups.Concat(beDups);
            File.WriteAllText(dup_jsonPath, JsonConvert.SerializeObject(new { total = dups.Count(), backend = beDups.Count(), frontend = feDups.Count(), result = dups }, Formatting.Indented));
        }
    }
}