using IronXL;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;

bool excelProcess = true;
bool jsonProcess = true;
bool jsonMetaProcess = true;
bool autoSizeRow = false;
bool autoSizeColumn = false;
int levelPrefix = 1;
bool autoDoublePrefixForDupKey = false;

string version = "v1";
string compareWithFileName = "en_flat";

string path = "C:\\Users\\Admin\\Desktop\\duplicate\\";
string prefixFileName = "7.0_en";

string i_excelPath = $"{path}{prefixFileName}_{version}.xlsx";
//string i_jsonPath = $"{path}{prefixFileName}_{version}.json";

string o_excelPath = $"{path}{prefixFileName}_unique_{version}.xlsx";
string o_jsonPath = $"{path}{prefixFileName}_unique_{version}.json";
string m_jsonPath = $"{path}{prefixFileName}_meta_{version}.json";

string c_excelPath = $"{path}{compareWithFileName}.xlsx";

string dup_jsonPath = $"{path}{prefixFileName}_dup_{version}.json";
string t_excelPath = $"{path}__temp__.xlsx";

string common_fe_prefix = "common";
string common_be_prefix = "common.backend";
string beKeyword = "backendService";

if (!excelProcess && !jsonProcess)
    return;

Console.WriteLine("Started");
Stopwatch stopwatch = new Stopwatch();
stopwatch.Start();

bool processCompare = !string.IsNullOrWhiteSpace(compareWithFileName);

var wb_original = WorkBook.Load(i_excelPath);
if(wb_original != null)
{
    if (File.Exists(t_excelPath))
        File.Delete(t_excelPath);

    var wbTemp = wb_original.SaveAs(t_excelPath);
    wb_original.Close();

    File.Delete(wbTemp.FilePath);

    var wsTemp = wbTemp.GetWorkSheet("en") ?? wbTemp.DefaultWorkSheet;

    int colValueIndex = 1; // en    
    string checkDupTargetColumn = wsTemp.Columns[colValueIndex].Rows[0].StringValue;
    var addressInCol = Regex.Replace(wsTemp.Columns[colValueIndex - 1].Rows[0].RangeAddressAsString, @"\d" , string.Empty);

    // keep current char case
    var originalTargetCol = wsTemp.Columns[colValueIndex]
                                  .Where(s => s.RowIndex != 0); // not count header row

    Console.WriteLine("total rows: " + originalTargetCol.Count());

    var uniqueListValues = originalTargetCol.GroupBy(g => g.Value.ToString());

    Console.WriteLine("total groups: " + uniqueListValues.Count());

    WorkBook? rWb = excelProcess ? WorkBook.Create(ExcelFileFormat.XLSX) : null;
    WorkSheet? rWs = null;
    if (rWb != null)
    {
        rWb.DefaultWorkSheet.Name = wsTemp.Name;
        rWs = rWb.DefaultWorkSheet;

        rWs["A1"].Value = "compare_status" + (processCompare ? $" (with {compareWithFileName})" : string.Empty);
        rWs["A1"].Style.Font.Bold = true;

        rWs["B1"].Value = "compare_logs";
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
    List<Item> compareDupResults = new List<Item>();

    List<IGrouping<string, Item>> dupkeyFe, dupkeyBe;

    Common.ProcessMetadata(autoDoublePrefixForDupKey, addressInCol, levelPrefix, common_fe_prefix, common_be_prefix, beKeyword, wsTemp, uniqueListValues, duplicatedFeKeyGroups, duplicatedBeKeyGroups, out dupkeyFe, out dupkeyBe);

    if (excelProcess)
    {
        rWs.AutoSizeRow(0);

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

                var targetCompareValueCol = cWs.Columns.FirstOrDefault(f => f.Rows[0].StringValue == checkDupTargetColumn);
                
                var compareUniqueListValues = targetCompareValueCol
                                                 .Where(s => s.RowIndex != 0) // not count header row
                                                 .GroupBy(g => g.Value.ToString());

                var rangeAddr = targetCompareValueCol.RangeAddressAsString;
                if (rangeAddr.Contains(":"))
                {
                    rangeAddr = Regex.Replace(rangeAddr.Split(':')[0], @"\d", string.Empty);                   
                }
                var compareAddressInCol = Common.ColumnIndexToColumnLetter(Common.ColumnLetterToColumnIndex(rangeAddr) - 1);

                Common.ProcessMetadata(autoDoublePrefixForDupKey, compareAddressInCol, levelPrefix, common_fe_prefix, common_be_prefix, beKeyword, cWs, compareUniqueListValues, compareFeKeyGroups, compareBeKeyGroups, out compareDupkeyFe, out compareDupkeyBe);

                Common.ProcessCompare(rWs, duplicatedFeKeyGroups, compareResults, compareFeKeyGroups);

                Common.ProcessCompare(rWs, dupkeyFe, compareDupResults, compareDupkeyFe);

                Common.ProcessCompare(rWs, dupkeyBe, compareDupResults, compareDupkeyBe);

                Common.ProcessCompare(rWs, duplicatedBeKeyGroups, compareResults, compareBeKeyGroups);
            }

            var _compareResults = compareResults.Concat(compareDupResults).OrderBy(o => (o.Key.Contains(".") ? o.Key.Split(".")[o.Key.Split(".").Length - 1] : o.Key))
                                       .ThenBy(t => t.Value)
                                       .ThenBy(t => t.CompareStatus)
                                       .ThenBy(t => t.CompareNote)
                                       .ToList();

            foreach (var cp in _compareResults)
            {
                rWs[$"B{i}"].Value = cp.CompareNote;
                rWs[$"A{i}"].Value = cp.CompareStatus.ToString();
                rWs[$"C{i}"].Value = cp.Key;
                rWs[$"D{i}"].Value = string.Join(Environment.NewLine, cp.DuplicatedKeys);
                rWs[$"E{i}"].Value = cp.Value;

                rWs[$"B{i}"].Style.WrapText = true;
                rWs[$"D{i}"].Style.WrapText = true;
                rWs[$"E{i}"].Style.WrapText = true;

                if (cp.CompareStatus == ItemStatus.ADDED_NEW)
                {
                    rWs[$"A{i}"].Style.Font.Color = "#05F50D";
                    rWs[$"C{i}:E{i}"].Style.Font.Color = "#05F50D";
                }
                else if (cp.CompareStatus == ItemStatus.REMOVED)
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
        }

        foreach (var fe in duplicatedFeKeyGroups)
        {
            rWs[$"C{i}"].Value = fe.Key;
            rWs[$"D{i}"].Value = string.Join(Environment.NewLine, fe.DuplicatedKeys);
            rWs[$"E{i}"].Value = fe.Value;

            rWs[$"B{i}"].Style.WrapText = true;
            rWs[$"D{i}"].Style.WrapText = true;
            rWs[$"E{i}"].Style.WrapText = true;

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

                    rWs[$"B{i}"].Style.WrapText = true;
                    rWs[$"D{i}"].Style.WrapText = true;
                    rWs[$"E{i}"].Style.WrapText = true;

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

                    rWs[$"B{i}"].Style.WrapText = true;
                    rWs[$"D{i}"].Style.WrapText = true;
                    rWs[$"E{i}"].Style.WrapText = true;

                    if (autoSizeRow)
                        rWs.AutoSizeRow(i - 1);

                    i++;
                }

        foreach (var be in duplicatedBeKeyGroups)
        {
            rWs[$"C{i}"].Value = be.Key;
            rWs[$"D{i}"].Value = string.Join(Environment.NewLine, be.DuplicatedKeys);
            rWs[$"E{i}"].Value = be.Value;

            rWs[$"B{i}"].Style.WrapText = true;
            rWs[$"D{i}"].Style.WrapText = true;
            rWs[$"E{i}"].Style.WrapText = true;

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
            rWs.AutoSizeColumn(4);
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

            if (jsonMetaProcess)
            {
                JObject metaResult = new JObject();

                foreach (var gr in duplicatedFeKeyGroups)
                    foreach (var item in gr.DuplicatedKeys)
                        metaResult.Add(item, gr.Key);

                foreach (var group in dupkeyFe)
                    foreach (var item in group)
                        foreach (var d in item.DuplicatedKeys)
                            if (!metaResult.ContainsKey(d))
                                metaResult.Add(d, d);
                            else if (metaResult[d].ToString() != item.Value.ToString())
                                Console.WriteLine(d + ": " + metaResult[d].ToString() + " - " + item.Value.ToString());

                foreach (var group in dupkeyBe)
                    foreach (var item in group)
                        foreach (var d in item.DuplicatedKeys)
                            if (!metaResult.ContainsKey(d))
                                metaResult.Add(d, d);
                            else if (metaResult[d].ToString() != item.Value.ToString())
                                Console.WriteLine(d + ": " + metaResult[d].ToString() + " - " + item.Value.ToString());

                foreach (var gr in duplicatedBeKeyGroups)
                    foreach (var item in gr.DuplicatedKeys)
                        metaResult.Add(item, gr.Key);

                File.WriteAllText(m_jsonPath, JsonConvert.SerializeObject(metaResult, Formatting.Indented));
            }

            foreach (var gr in compareResults.Where(w => w.CompareStatus != ItemStatus.REMOVED))
                combines.Add(gr.Key, gr.Value.ToString());

            foreach (var gr in compareDupResults.Where(w => w.CompareStatus != ItemStatus.REMOVED))
                foreach (var d in gr.DuplicatedKeys)
                    combines.Add(d, gr.Value.ToString());

            foreach (var gr in duplicatedFeKeyGroups)
                combines.Add(gr.Key, gr.Value.ToString());

            foreach (var group in dupkeyFe)
                foreach (var item in group)
                    foreach (var d in item.DuplicatedKeys)
                        if (!combines.ContainsKey(d))
                            combines.Add(d, item.Value.ToString());
                        else if (combines[d].ToString() != item.Value.ToString())
                            Console.WriteLine(d + ": " + combines[d].ToString() + " - " + item.Value.ToString());

            foreach (var group in dupkeyBe)
                foreach (var item in group)
                    foreach (var d in item.DuplicatedKeys)
                        if (!combines.ContainsKey(d))
                            combines.Add(d, item.Value.ToString());
                        else if (combines[d].ToString() != item.Value.ToString())
                            Console.WriteLine(d + ": " + combines[d].ToString() + " - " + item.Value.ToString());

            foreach (var gr in duplicatedBeKeyGroups)
                combines.Add(gr.Key, gr.Value.ToString());

            File.WriteAllText(o_jsonPath, JsonConvert.SerializeObject(combines, Formatting.Indented));
            var beDups = duplicatedBeKeyGroups.SelectMany(k => k.DuplicatedKeys);
            var feDups = duplicatedFeKeyGroups.SelectMany(s => s.DuplicatedKeys);
            var dups = feDups.Concat(beDups);
            File.WriteAllText(dup_jsonPath, JsonConvert.SerializeObject(new { total = dups.Count(), backend = beDups.Count(), frontend = feDups.Count(), result = dups }, Formatting.Indented));
        }
    }
    stopwatch.Stop();
    var tracking = stopwatch.Elapsed;
    Console.WriteLine("Finished in " + tracking.ToString());
}