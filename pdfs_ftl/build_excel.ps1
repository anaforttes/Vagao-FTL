Add-Type -AssemblyName System.Web.Extensions
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml

$code = @"
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Xml;

public class OcrWord { public string T { get; set; } public double X { get; set; } public double Y { get; set; } }
public class PdfOcr { public string File { get; set; } public double W { get; set; } public double H { get; set; } public List<OcrWord> Words { get; set; } }

public class ApiWagon {
    public int id { get; set; }
    public string codigo_vagao { get; set; }
    public string serie_vagao { get; set; }
    public string caminho_arquivo { get; set; }
    public string nome_arquivo { get; set; }
    public string data_insercao { get; set; }
}
public class ApiResponse {
    public List<ApiWagon> content { get; set; }
}

public class WagonExtra {
    public string InicioValidade { get; set; }
    public string FimValidade { get; set; }
    public int ReguaMin { get; set; }
    public int ReguaMax { get; set; }
    public string Situacao { get; set; }
    public string DataInsercao { get; set; }
    public int[] DomoVol { get; set; }
    public int[] DomoMm { get; set; }
    public List<int> TanqueVol { get; set; }
    public List<int> TanqueMm { get; set; }
}

public class NewExcelUpdater {
    public static string ParseDate(string text) {
        if (string.IsNullOrEmpty(text)) return null;
        string t = text.Trim().Replace("O","0").Replace("o","0").Replace("I","1").Replace("l","1").Replace("|","1").Replace("!","1");
        t = Regex.Replace(t, @"[^\d/.\-]", "");
        Match m;
        m = Regex.Match(t, @"^(\d{2})[/.\-](\d{2})[/.\-](20\d{2}|19\d{2})$");
        if (m.Success) return m.Groups[1].Value + "/" + m.Groups[2].Value + "/" + m.Groups[3].Value;
        m = Regex.Match(t, @"^(\d{2})1(\d{2})[/.\-](20\d{2}|19\d{2})$");
        if (m.Success) return m.Groups[1].Value + "/" + m.Groups[2].Value + "/" + m.Groups[3].Value;
        m = Regex.Match(t, @"^(\d{2})[/.\-](\d{2})1(20\d{2}|19\d{2})$");
        if (m.Success) return m.Groups[1].Value + "/" + m.Groups[2].Value + "/" + m.Groups[3].Value;
        m = Regex.Match(t, @"^(\d{2})1(\d{2})1(20\d{2}|19\d{2})$");
        if (m.Success) return m.Groups[1].Value + "/" + m.Groups[2].Value + "/" + m.Groups[3].Value;
        m = Regex.Match(t, @"^(\d{2})(\d{2})(20\d{2}|19\d{2})$");
        if (m.Success) return m.Groups[1].Value + "/" + m.Groups[2].Value + "/" + m.Groups[3].Value;
        return null;
    }

    public static string GetColName(int colIndex1Based) {
        int dividend = colIndex1Based;
        string columnName = String.Empty;
        while (dividend > 0) {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
        }
        return columnName;
    }

    public static int DateToExcelSerial(string ddMMyyyy) {
        DateTime dt;
        if (DateTime.TryParseExact(ddMMyyyy, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) {
            DateTime baseDate = new DateTime(1899, 12, 30);
            return (int)(dt - baseDate).TotalDays;
        }
        return 0;
    }

    public static void UpdateNewFile(string apiJsonPath, string ocrPath, string targetXlsx) {
        var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        var apiResp = ser.Deserialize<ApiResponse>(File.ReadAllText(apiJsonPath));
        var ocrList = ser.Deserialize<List<PdfOcr>>(File.ReadAllText(ocrPath));
        var ocrByFile = ocrList.ToDictionary(x => x.File, x => x);

        var apiByCode = new Dictionary<string, ApiWagon>();
        foreach (var w in apiResp.content) {
            string clean = Regex.Replace(w.codigo_vagao ?? "", @"[^0-9]", "");
            if (!string.IsNullOrEmpty(clean)) {
                apiByCode[clean] = w;
                apiByCode[clean.TrimStart('0')] = w;
            }
        }

        var iniOverrides = new Dictionary<string, string> {
            {"6089518_TCC608951-8-AF-10-12-25.pdf", "28/11/2025"},
            {"6091555_TCC609155-5-AF-19-12-22.pdf", "19/12/2022"},
            {"6091253_TCC6091253.pdf", "13/11/2018"},
            {"6279112_TCC627911-2-AF-08-12-25.pdf", "17/11/2025"},
            {"6089003_TCC608900-3-AF-10-12-25.pdf", "26/11/2025"},
            {"6091695_TCC609169-5-AF-06-11-25.pdf", "12/08/2025"},
            {"6007007_TCC600700-7-AF-07-12-25.pdf", "18/11/2025"},
            {"6088872_TCC608887-2-AF-16-09-25.pdf", "26/07/2025"},
            {"6175953_TCC617595-3-AF-21-09-22.pdf", "21/09/2022"},
            {"6090061_TCC609006-1-AF-18-06-25.pdf", "21/05/2025"},
            {"6090672_TCC609067-2-AF-19-10-22.pdf", "19/10/2022"},
            {"6006035_TCC600603-5-AF-09-12-25.pdf", "24/11/2025"},
            {"6088805_TCC608880-5-AF-07-12-25.pdf", "18/11/2025"},
            {"6091474_TCC609147-4-AF-18-10-22.pdf", "18/10/2022"},
            {"6175449_TCC617544-9-AF-19-10-22.pdf", "19/10/2022"},
            {"6021883_TCC602188-3-AF-16-11-22.pdf", "16/11/2022"},
            {"6091431_TCC609143-1-AF-07-12-25.pdf", "19/11/2025"},
            {"6090958_TCC609095-8-AF-07-12-25.pdf", "18/11/2025"},
            {"6279881_TCC627988-1-AF-07-12-25.pdf", "18/11/2025"},
            {"6089194_TCC608919-4-AF-09-12-25.pdf", "25/11/2025"},
            {"6089551_TCC608955-1-AF-06-11-25.pdf", "09/09/2025"},
            {"6279431_TCC627943-1-AF-19-10-22.pdf", "19/10/2022"},
            {"6090737_TCC609073-7-AF-07-12-25.pdf", "18/11/2025"},
            {"6090141_TCC609014-1-AF-17-08-26.pdf", "05/08/2026"}
        };

        var fimOverrides = new Dictionary<string, string> {
            {"6089518_TCC608951-8-AF-10-12-25.pdf", "10/12/2029"},
            {"6091555_TCC609155-5-AF-19-12-22.pdf", "18/12/2026"},
            {"6091253_TCC6091253.pdf", "15/11/2022"},
            {"6090141_TCC609014-1-AF-17-08-26.pdf", "17/08/2030"}
        };

        var domoMmOverrides = new Dictionary<string, int[]> {
            {"6279112_TCC627911-2-AF-08-12-25.pdf", new[]{513,492,471,450}},
            {"6089003_TCC608900-3-AF-10-12-25.pdf", new[]{491,469,447,426}},
            {"6091695_TCC609169-5-AF-06-11-25.pdf", new[]{479,459,438,418}},
            {"6007007_TCC600700-7-AF-07-12-25.pdf", new[]{506,484,462,440}},
            {"6088872_TCC608887-2-AF-16-09-25.pdf", new[]{489,467,443,422}},
            {"6091474_TCC609147-4-AF-18-10-22.pdf", new[]{469,448,427,405}},
            {"6090958_TCC609095-8-AF-07-12-25.pdf", new[]{498,477,465,434}},
            {"6089194_TCC608919-4-AF-09-12-25.pdf", new[]{489,467,446,424}},
            {"6089232_TCC608923-2-AF-20-06-25.pdf", new[]{482,460,438,417}},
            {"6279121_TCC627912-1-AF-30-07-26.pdf", new[]{489,467,445,423}},
            {"6090141_TCC609014-1-AF-17-08-26.pdf", new[]{488,466,443,421}}
        };

        var tanqueMmOverrides = new Dictionary<string, int[]> {
            {"6089020_TCC608902-0-AF-20-10-22.pdf", new[]{510,521,546,560,572,584,595}},
            {"6088872_TCC608887-2-AF-16-09-25.pdf", new[]{511,532,548,562,574,585,595,604}},
            {"6088759_TCC608875-9-AF-17-11-22.pdf", new[]{533,548,561,574,587,597,608,618}},
            {"6089283_TCC608928-3-AF-08-12-25.pdf", new[]{573,591,605,618,631,641,650,659}},
            {"6090346_TCC609034-6-AF-16-09-25.pdf", new[]{504,522,540,557,568,580,591}},
            {"6090613_TCC609061-3-AF-07-12-25.pdf", new[]{537,549,564,577,589,601,612,622}},
            {"6279384_TCC627938-4-AF-17-06-25.pdf", new[]{509,524,540,556,567,578,588}},
            {"6089771_TCC608977-1-AF-18-06-25.pdf", new[]{508,526,541,554,567,578,587,596}},
            {"6090141_TCC609014-1-AF-17-08-26.pdf", new[]{509,529,547,560,572,583,593,602}}
        };

        DateTime today = new DateTime(2026, 11, 9);

        Func<string, int, WagonExtra> getExtra = (veicCode, capNom) => {
            string clean = Regex.Replace(veicCode ?? "", @"[^0-9]", "");
            ApiWagon apiW = null;
            if (!apiByCode.TryGetValue(clean, out apiW)) apiByCode.TryGetValue(clean.TrimStart('0'), out apiW);
            string fname = apiW.codigo_vagao + "_" + apiW.nome_arquivo;
            List<OcrWord> words = ocrByFile.ContainsKey(fname) ? (ocrByFile[fname].Words ?? new List<OcrWord>()) : new List<OcrWord>();

            string fimVal = fimOverrides.ContainsKey(fname) ? fimOverrides[fname] : null;
            if (fimVal == null) {
                var topR = words.Where(w => w.X > 0.55 && w.Y < 0.25).OrderBy(w => w.Y).ThenBy(w => w.X).ToList();
                foreach (var w in topR) { fimVal = ParseDate(w.T); if (fimVal != null) break; }
                if (fimVal == null) {
                    for (int i = 0; i < topR.Count - 1; i++) {
                        if (Math.Abs(topR[i].Y - topR[i+1].Y) < 0.01) {
                            fimVal = ParseDate(topR[i].T + topR[i+1].T);
                            if (fimVal != null) break;
                        }
                    }
                }
            }

            string iniVal = iniOverrides.ContainsKey(fname) ? iniOverrides[fname] : null;
            if (iniVal == null) {
                var botL = words.Where(w => w.X < 0.45 && w.Y > 0.75).OrderBy(w => w.Y).ThenBy(w => w.X).ToList();
                foreach (var w in botL) { iniVal = ParseDate(w.T); if (iniVal != null) break; }
                if (iniVal == null) {
                    for (int i = 0; i < botL.Count - 1; i++) {
                        if (Math.Abs(botL[i].Y - botL[i+1].Y) < 0.01) {
                            iniVal = ParseDate(botL[i].T + botL[i+1].T);
                            if (iniVal != null) break;
                        }
                    }
                }
            }

            var medFinais = words.Where(w => Regex.IsMatch(w.T, "FINAIS|CHEIO|NOMINAL") && w.Y > 0.55 && w.Y < 0.80).OrderBy(w => w.Y).FirstOrDefault();
            double yMax = medFinais != null ? medFinais.Y - 0.003 : 0.64;
            var tHeader = words.Where(w => w.X < 0.50 && Regex.IsMatch(w.T, "VOLUME|VAZIO|QUANT") && w.Y > 0.44 && w.Y < 0.56).ToList();
            double yMinT = tHeader.Count > 0 ? tHeader.Min(w => w.Y) + 0.005 : 0.49;

            double yRow0 = yMinT + 0.012;
            var tMmWords = words.Where(w => w.Y > yMinT && w.Y < yMax && w.X > 0.35 && w.X < 0.49)
                .Select(w => new { W = w, ValStr = Regex.Replace(w.T.Replace("O","0").Replace("o","0").Replace("I","1").Replace("l","1").Replace("B","8"), @"[^0-9]", "") })
                .Where(x => Regex.IsMatch(x.ValStr, @"^\d{3,4}$"))
                .OrderBy(x => x.W.Y).ToList();
            if (tMmWords.Count > 0) yRow0 = tMmWords[0].W.Y;

            List<int> tMms = tanqueMmOverrides.ContainsKey(fname)
                ? tanqueMmOverrides[fname].ToList()
                : tMmWords.Select(x => int.Parse(x.ValStr)).ToList();

            var tVols = new List<int>();
            if (tMms.Count > 0) {
                tVols.Add(capNom);
                int firstHundred = ((capNom - 1) / 100) * 100;
                for (int k = 1; k < tMms.Count; k++) tVols.Add(firstHundred - (k - 1) * 100);
            }

            int[] dVols = new int[4] { capNom + 50, capNom + 100, capNom + 150, capNom + 200 };
            int[] dMms = new int[4];
            if (domoMmOverrides.ContainsKey(fname)) {
                dMms = domoMmOverrides[fname];
            } else {
                var dHeader = words.Where(w => w.X > 0.50 && Regex.IsMatch(w.T, "VOLUME|VAZIO|QUANT") && w.Y > 0.44 && w.Y < 0.56).ToList();
                double yMinD = dHeader.Count > 0 ? dHeader.Min(w => w.Y) + 0.005 : 0.49;
                var dWords = words.Where(w => w.Y > yMinD && w.Y < yMax && w.X > 0.70 && w.X < 0.90)
                    .Select(w => new { W = w, ValStr = Regex.Replace(w.T.Replace("O","0").Replace("o","0").Replace("I","1").Replace("l","1").Replace("B","8"), @"[^0-9]", "") })
                    .Where(x => Regex.IsMatch(x.ValStr, @"^\d{3,4}$"))
                    .OrderBy(x => x.W.Y).ToList();

                if (dWords.Count >= 4) {
                    for (int i = 0; i < 4; i++) dMms[i] = int.Parse(dWords[i].ValStr);
                } else if (dWords.Count == 3) {
                    double yA = dWords[0].W.Y, yB = dWords[1].W.Y, yC = dWords[2].W.Y;
                    int vA = int.Parse(dWords[0].ValStr), vB = int.Parse(dWords[1].ValStr), vC = int.Parse(dWords[2].ValStr);
                    double dy1 = yB - yA, dy2 = yC - yB;
                    if (dy1 > 0.020) { dMms[0] = vA; dMms[1] = (int)Math.Round((vA + vB) / 2.0); dMms[2] = vB; dMms[3] = vC; }
                    else if (dy2 > 0.020) { dMms[0] = vA; dMms[1] = vB; dMms[2] = (int)Math.Round((vB + vC) / 2.0); dMms[3] = vC; }
                    else {
                        int step = (int)Math.Round((vA - vC) / 2.0);
                        if (yA - yRow0 > 0.009) { dMms[0] = vA + step; dMms[1] = vA; dMms[2] = vB; dMms[3] = vC; }
                        else { dMms[0] = vA; dMms[1] = vB; dMms[2] = vC; dMms[3] = vC - step; }
                    }
                }
            }

            string situacao = "Regular";
            DateTime dtFim;
            if (DateTime.TryParseExact(fimVal, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtFim)) {
                if (dtFim < today) situacao = "Vencido";
            }

            string dataIns = apiW.data_insercao ?? "";
            if (Regex.IsMatch(dataIns, @"^\d{4}-\d{2}-\d{2}")) {
                var parts = dataIns.Substring(0, 10).Split('-');
                dataIns = parts[2] + "/" + parts[1] + "/" + parts[0];
            }

            return new WagonExtra {
                InicioValidade = iniVal ?? "",
                FimValidade = fimVal ?? "",
                ReguaMin = tMms.Count > 0 ? tMms[0] : 0,
                ReguaMax = tMms.Count > 0 ? tMms[tMms.Count - 1] : 0,
                Situacao = situacao,
                DataInsercao = dataIns,
                DomoVol = dVols,
                DomoMm = dMms,
                TanqueVol = tVols,
                TanqueMm = tMms
            };
        };

        using (var archive = ZipFile.Open(targetXlsx, ZipArchiveMode.Update)) {
            var ssEntry = archive.GetEntry("xl/sharedStrings.xml");
            var ssDoc = new XmlDocument();
            using (var s = ssEntry.Open()) { ssDoc.Load(s); }
            var nsMgr = new XmlNamespaceManager(ssDoc.NameTable);
            string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            nsMgr.AddNamespace("x", ns);

            var sharedStrings = new List<string>();
            var sharedMap = new Dictionary<string, int>();
            foreach (XmlNode si in ssDoc.DocumentElement.SelectNodes("x:si", nsMgr)) {
                string val = si.InnerText;
                if (!sharedMap.ContainsKey(val)) sharedMap[val] = sharedStrings.Count;
                sharedStrings.Add(val);
            }

            Func<string, int> getOrAddString = (str) => {
                if (str == null) str = "";
                int idx;
                if (sharedMap.TryGetValue(str, out idx)) return idx;
                idx = sharedStrings.Count;
                sharedStrings.Add(str);
                sharedMap[str] = idx;
                var siNode = ssDoc.CreateElement("si", ns);
                var tNode = ssDoc.CreateElement("t", ns);
                tNode.InnerText = str;
                siNode.AppendChild(tNode);
                ssDoc.DocumentElement.AppendChild(siNode);
                return idx;
            };

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
            var sheetDoc = new XmlDocument();
            using (var s = sheetEntry.Open()) { sheetDoc.Load(s); }

            var sheetData = sheetDoc.DocumentElement.SelectSingleNode("x:sheetData", nsMgr);
            var rows = sheetData.SelectNodes("x:row", nsMgr);

            Action<XmlElement, int, int, int, string, bool> setCell = (rowEl, rowIdx, colIdx, styleIdx, val, isString) => {
                string cellRef = GetColName(colIdx) + rowIdx.ToString();
                var existing = rowEl.SelectSingleNode("x:c[@r='" + cellRef + "']", nsMgr) as XmlElement;
                if (existing != null) rowEl.RemoveChild(existing);
                var cEl = sheetDoc.CreateElement("c", ns);
                cEl.SetAttribute("r", cellRef);
                cEl.SetAttribute("s", styleIdx.ToString());
                if (!string.IsNullOrEmpty(val)) {
                    if (isString) {
                        cEl.SetAttribute("t", "s");
                        var vEl = sheetDoc.CreateElement("v", ns);
                        vEl.InnerText = getOrAddString(val).ToString();
                        cEl.AppendChild(vEl);
                    } else {
                        var vEl = sheetDoc.CreateElement("v", ns);
                        vEl.InnerText = val;
                        cEl.AppendChild(vEl);
                    }
                }
                rowEl.AppendChild(cEl);
            };

            // Header row (Row 1): keep A1..I1, add J1..AR1
            var headerRow = rows[0] as XmlElement;
            headerRow.SetAttribute("spans", "1:44");
            string[] newHeaders = new string[] {
                "Início Validade",
                "Régua Mínima (mm)",
                "Régua Máxima (mm)",
                "Situação",
                "Data de Inserção",
                "Domo 4 (L)", "Domo 4 (mm)",
                "Domo 3 (L)", "Domo 3 (mm)",
                "Domo 2 (L)", "Domo 2 (mm)",
                "Domo 1 (L)", "Domo 1 (mm)",
                "Tanque 0 (L)", "Tanque 0 (mm)",
                "Tanque 1 (L)", "Tanque 1 (mm)",
                "Tanque 2 (L)", "Tanque 2 (mm)",
                "Tanque 3 (L)", "Tanque 3 (mm)",
                "Tanque 4 (L)", "Tanque 4 (mm)",
                "Tanque 5 (L)", "Tanque 5 (mm)",
                "Tanque 6 (L)", "Tanque 6 (mm)",
                "Tanque 7 (L)", "Tanque 7 (mm)",
                "Tanque 8 (L)", "Tanque 8 (mm)",
                "Tanque 9 (L)", "Tanque 9 (mm)",
                "Tanque 10 (L)", "Tanque 10 (mm)"
            };
            for (int i = 0; i < newHeaders.Length; i++) {
                int colIdx = 10 + i; // J = 10
                int st = (i == 0 || i == 4) ? 5 : 3;
                setCell(headerRow, 1, colIdx, st, newHeaders[i], true);
            }

            // Data rows (Row 2..370)
            for (int rIdx = 1; rIdx < rows.Count; rIdx++) {
                var rowEl = rows[rIdx] as XmlElement;
                int rowNum = int.Parse(rowEl.Attributes["r"].Value);
                rowEl.SetAttribute("spans", "1:44");

                var bCell = rowEl.SelectSingleNode("x:c[@r='B" + rowNum + "']", nsMgr) as XmlElement;
                var eCell = rowEl.SelectSingleNode("x:c[@r='E" + rowNum + "']", nsMgr) as XmlElement;
                string veic = bCell.InnerText;
                if (bCell.GetAttribute("t") == "s") veic = sharedStrings[int.Parse(veic)];
                string capStr = eCell.InnerText;
                if (eCell.GetAttribute("t") == "s") capStr = sharedStrings[int.Parse(capStr)];
                int capNom = int.Parse(capStr);

                var ext = getExtra(veic, capNom);

                int serialIni = DateToExcelSerial(ext.InicioValidade);
                int serialIns = DateToExcelSerial(ext.DataInsercao);

                if (serialIni > 0) setCell(rowEl, rowNum, 10, 6, serialIni.ToString(), false);
                else setCell(rowEl, rowNum, 10, 2, ext.InicioValidade, true);

                setCell(rowEl, rowNum, 11, 2, ext.ReguaMin.ToString(), false);
                setCell(rowEl, rowNum, 12, 2, ext.ReguaMax.ToString(), false);
                setCell(rowEl, rowNum, 13, 2, ext.Situacao, true);

                if (serialIns > 0) setCell(rowEl, rowNum, 14, 6, serialIns.ToString(), false);
                else setCell(rowEl, rowNum, 14, 2, ext.DataInsercao, true);

                // Domo 4, 3, 2, 1 => cols 15..22 (O..V)
                for (int d = 3; d >= 0; d--) {
                    int colL = 15 + (3 - d) * 2;
                    int colMm = colL + 1;
                    setCell(rowEl, rowNum, colL, 2, ext.DomoVol[d].ToString(), false);
                    setCell(rowEl, rowNum, colMm, 2, ext.DomoMm[d].ToString(), false);
                }

                // Tanque 0..10 => cols 23..44 (W..AR)
                for (int t = 0; t <= 10; t++) {
                    int colL = 23 + t * 2;
                    int colMm = colL + 1;
                    if (t < ext.TanqueVol.Count) {
                        setCell(rowEl, rowNum, colL, 2, ext.TanqueVol[t].ToString(), false);
                        setCell(rowEl, rowNum, colMm, 2, ext.TanqueMm[t].ToString(), false);
                    } else {
                        setCell(rowEl, rowNum, colL, 2, "", false);
                        setCell(rowEl, rowNum, colMm, 2, "", false);
                    }
                }
            }

            var dimNode = sheetDoc.DocumentElement.SelectSingleNode("x:dimension", nsMgr) as XmlElement;
            if (dimNode != null) dimNode.SetAttribute("ref", "A1:AR" + rows.Count.ToString());

            var afNode = sheetDoc.DocumentElement.SelectSingleNode("x:autoFilter", nsMgr) as XmlElement;
            if (afNode != null) afNode.SetAttribute("ref", "A1:AR1");

            var colsNode = sheetDoc.DocumentElement.SelectSingleNode("x:cols", nsMgr) as XmlElement;
            if (colsNode != null) {
                colsNode.RemoveAll();
                Action<int, int, double, int, bool> addColDef = (minC, maxC, w, st, hid) => {
                    var c = sheetDoc.CreateElement("col", ns);
                    c.SetAttribute("min", minC.ToString());
                    c.SetAttribute("max", maxC.ToString());
                    c.SetAttribute("width", w.ToString(CultureInfo.InvariantCulture));
                    c.SetAttribute("style", st.ToString());
                    if (hid) c.SetAttribute("hidden", "1");
                    c.SetAttribute("customWidth", "1");
                    colsNode.AppendChild(c);
                };
                addColDef(1, 1, 7.3, 7, false);
                addColDef(2, 2, 10.0, 7, false);
                addColDef(3, 3, 7.8, 7, true);
                addColDef(4, 4, 8.8, 7, true);
                addColDef(5, 5, 15.2, 8, false);
                addColDef(6, 6, 17.8, 7, false);
                addColDef(7, 7, 18.0, 7, false);
                addColDef(8, 8, 16.0, 7, false);
                addColDef(9, 10, 15.0, 9, false);
                addColDef(11, 13, 14.0, 7, false);
                addColDef(14, 14, 15.0, 9, false);
                addColDef(15, 44, 12.5, 7, false);
            }

            ssDoc.DocumentElement.SetAttribute("count", sharedStrings.Count.ToString());
            ssDoc.DocumentElement.SetAttribute("uniqueCount", sharedStrings.Count.ToString());

            ssEntry.Delete();
            var newSsEntry = archive.CreateEntry("xl/sharedStrings.xml");
            using (var s = newSsEntry.Open()) { ssDoc.Save(s); }

            sheetEntry.Delete();
            var newSheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using (var s = newSheetEntry.Open()) { sheetDoc.Save(s); }
        }
        Console.WriteLine("Successfully updated: " + targetXlsx);
    }
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies System.Web.Extensions, System.IO.Compression, System.IO.Compression.FileSystem, System.Xml
if (-not (Test-Path "pdfs_ftl\VTs FTL NEW_backup.xlsx")) {
    Copy-Item "VTs FTL NEW.xlsx" "pdfs_ftl\VTs FTL NEW_backup.xlsx"
}
[NewExcelUpdater]::UpdateNewFile(
    (Resolve-Path "pdfs_ftl\vagoes_api.json").Path,
    (Resolve-Path "pdfs_ftl\ocr_words_368.json").Path,
    (Resolve-Path "VTs FTL NEW.xlsx").Path
)
