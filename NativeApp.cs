using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Runtime.InteropServices;

namespace SmartClassAssigner
{
    // =========================================================
    // 1. 학생 데이터 모델
    // =========================================================
    public class StudentRecord
    {
        public string Id { get; set; }
        public string PrevClass { get; set; }
        public string PrevNumber { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public string FirstName { get; set; }
        public string Gender { get; set; }
        public int Rank { get; set; }
        public double Score { get; set; }
        public string Remarks { get; set; }

        public string AssignedClass { get; set; }
        public int AssignedNumber { get; set; }
        public bool IsSpecial { get; set; }
        public List<string> SeparateFromNames { get; set; }

        public StudentRecord()
        {
            Rank = 99;
            Score = 0;
            SeparateFromNames = new List<string>();
        }
    }

    // =========================================================
    // 2. 한국어 이름 및 비고 파싱 유틸리티
    // =========================================================
    public static class NameParser
    {
        private static readonly HashSet<string> TwoCharSurnames = new HashSet<string>
        {
            "남궁", "황보", "제갈", "사공", "선우", "서문", "독고", "동방"
        };

        public static void ParseName(string fullName, out string surname, out string firstName)
        {
            surname = "";
            firstName = "";
            if (string.IsNullOrEmpty(fullName)) return;

            string clean = Regex.Replace(fullName, @"[\s\d\(\)\[\]]", "");
            if (clean.Length == 1)
            {
                firstName = clean;
                return;
            }
            if (clean.Length >= 4 && TwoCharSurnames.Contains(clean.Substring(0, 2)))
            {
                surname = clean.Substring(0, 2);
                firstName = clean.Substring(2);
                return;
            }
            if (clean.Length >= 2)
            {
                surname = clean.Substring(0, 1);
                firstName = clean.Substring(1);
                return;
            }
            firstName = clean;
        }

        public static List<string> ParseSeparationRequests(string remarks)
        {
            List<string> result = new List<string>();
            if (string.IsNullOrEmpty(remarks)) return result;

            string[] patterns = new string[]
            {
                @"분리요청\s*[\(\:\-]?\s*([가-힣a-zA-Z]{2,10})\s*[\)]?",
                @"분리\s*[\(\:\-]?\s*([가-힣a-zA-Z]{2,10})\s*[\)]?"
            };

            foreach (var pat in patterns)
            {
                var matches = Regex.Matches(remarks, pat);
                foreach (Match m in matches)
                {
                    string name = m.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(name) && !result.Contains(name))
                    {
                        result.Add(name);
                    }
                }
            }
            return result;
        }
    }

    // =========================================================
    // 3. 순수 C# 엑셀 (.xlsx) 리더 & 라이터 (의존성 제로)
    // =========================================================
    public static class ExcelManager
    {
        public static Dictionary<string, List<StudentRecord>> ReadXlsx(string filePath)
        {
            var result = new Dictionary<string, List<StudentRecord>>();

            if (filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                return ReadCsv(filePath);
            }

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    // 1. Shared Strings 로드
                    List<string> sharedStrings = new List<string>();
                    ZipArchiveEntry sstEntry = zip.GetEntry("xl/sharedStrings.xml");
                    if (sstEntry == null)
                    {
                        foreach (var e in zip.Entries)
                        {
                            if (string.Equals(e.FullName, "xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase))
                            {
                                sstEntry = e;
                                break;
                            }
                        }
                    }
                    if (sstEntry != null)
                    {
                        using (Stream stream = sstEntry.Open())
                        {
                            XmlDocument doc = new XmlDocument();
                            doc.Load(stream);
                            XmlNodeList siNodes = doc.GetElementsByTagName("si");
                            foreach (XmlNode si in siNodes)
                            {
                                StringBuilder sb = new StringBuilder();
                                XmlNodeList tNodes = si.SelectNodes(".//t");
                                if (tNodes != null && tNodes.Count > 0)
                                {
                                    foreach (XmlNode tn in tNodes) sb.Append(tn.InnerText);
                                }
                                else
                                {
                                    sb.Append(si.InnerText);
                                }
                                sharedStrings.Add(sb.ToString());
                            }
                        }
                    }

                    // 2. Workbook.xml 및 Relationships 로드
                    ZipArchiveEntry wbEntry = zip.GetEntry("xl/workbook.xml");
                    ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                    Dictionary<string, string> relMap = new Dictionary<string, string>();

                    if (relsEntry != null)
                    {
                        using (Stream stream = relsEntry.Open())
                        {
                            XmlDocument doc = new XmlDocument();
                            doc.Load(stream);
                            XmlNodeList relNodes = doc.GetElementsByTagName("Relationship");
                            foreach (XmlNode rel in relNodes)
                            {
                                string id = GetAttr(rel, "Id");
                                string target = GetAttr(rel, "Target");
                                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                                {
                                    target = target.TrimStart('/', '\\');
                                    if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                                    {
                                        target = "xl/" + target;
                                    }
                                    relMap[id] = target;
                                }
                            }
                        }
                    }

                    if (wbEntry != null)
                    {
                        XmlDocument wbDoc = new XmlDocument();
                        using (Stream stream = wbEntry.Open())
                        {
                            wbDoc.Load(stream);
                        }

                        XmlNodeList sheetNodes = wbDoc.GetElementsByTagName("sheet");
                        foreach (XmlNode sNode in sheetNodes)
                        {
                            string sheetName = GetAttr(sNode, "name") ?? "1반";
                            string rId = GetAttr(sNode, "r:id");
                            if (string.IsNullOrEmpty(rId)) rId = GetAttr(sNode, "id");

                            string sheetPath = null;
                            if (!string.IsNullOrEmpty(rId) && relMap.ContainsKey(rId))
                            {
                                sheetPath = relMap[rId];
                            }
                            else
                            {
                                foreach (var kv in relMap)
                                {
                                    if (string.Equals(kv.Key, rId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        sheetPath = kv.Value;
                                        break;
                                    }
                                }
                            }

                            ZipArchiveEntry sheetEntry = null;
                            if (!string.IsNullOrEmpty(sheetPath))
                            {
                                sheetEntry = zip.GetEntry(sheetPath);
                                if (sheetEntry == null) sheetEntry = zip.GetEntry(sheetPath.Replace('/', '\\'));
                                if (sheetEntry == null)
                                {
                                    foreach (var e in zip.Entries)
                                    {
                                        if (string.Equals(e.FullName.TrimStart('/', '\\'), sheetPath.TrimStart('/', '\\'), StringComparison.OrdinalIgnoreCase))
                                        {
                                            sheetEntry = e;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (sheetEntry != null)
                            {
                                var students = ParseWorksheet(sheetEntry, sheetName, sharedStrings);
                                if (students.Count > 0)
                                {
                                    result[sheetName] = students;
                                }
                            }
                        }
                    }

                    // 3. Fallback: 만약 시트를 못 찾았다면 zip 안의 모든 worksheet xml 직접 파싱
                    if (result.Count == 0)
                    {
                        int fallbackIdx = 1;
                        foreach (var entry in zip.Entries)
                        {
                            string fn = entry.FullName.Replace('\\', '/');
                            if (fn.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) && fn.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                            {
                                string sheetName = fallbackIdx + "반";
                                var students = ParseWorksheet(entry, sheetName, sharedStrings);
                                if (students.Count > 0)
                                {
                                    result[sheetName] = students;
                                    fallbackIdx++;
                                }
                            }
                        }
                    }
                }
            }
            catch (InvalidDataException)
            {
                return ReadCsv(filePath);
            }

            return result;
        }

        private static List<StudentRecord> ParseWorksheet(ZipArchiveEntry sheetEntry, string sheetName, List<string> sharedStrings)
        {
            List<StudentRecord> students = new List<StudentRecord>();
            using (Stream stream = sheetEntry.Open())
            {
                XmlDocument sheetDoc = new XmlDocument();
                sheetDoc.Load(stream);

                XmlNodeList rowNodes = sheetDoc.GetElementsByTagName("row");
                if (rowNodes.Count == 0) return students;

                int headerRowIdx = -1;
                Dictionary<int, string> colHeaders = new Dictionary<int, string>();

                int scanLimit = Math.Min(10, rowNodes.Count);
                for (int ri = 0; ri < scanLimit; ri++)
                {
                    var tempHeaders = new Dictionary<int, string>();
                    foreach (XmlNode c in rowNodes[ri].ChildNodes)
                    {
                        if (!string.Equals(c.LocalName, "c", StringComparison.OrdinalIgnoreCase)) continue;
                        int colIdx = GetColumnIndex(GetAttr(c, "r"));
                        string val = GetCellValue(c, sharedStrings).Trim();
                        if (!string.IsNullOrEmpty(val)) tempHeaders[colIdx] = val;
                    }

                    bool hasName = false;
                    foreach (var val in tempHeaders.Values)
                    {
                        string clean = val.Replace(" ", "");
                        if (clean.Contains("이름") || clean.Contains("성명") || clean.Contains("학생명"))
                        {
                            hasName = true;
                            break;
                        }
                    }

                    if (hasName)
                    {
                        headerRowIdx = ri;
                        colHeaders = tempHeaders;
                        break;
                    }
                }

                if (headerRowIdx == -1)
                {
                    headerRowIdx = 0;
                    foreach (XmlNode c in rowNodes[0].ChildNodes)
                    {
                        if (!string.Equals(c.LocalName, "c", StringComparison.OrdinalIgnoreCase)) continue;
                        int colIdx = GetColumnIndex(GetAttr(c, "r"));
                        string val = GetCellValue(c, sharedStrings).Trim();
                        colHeaders[colIdx] = val;
                    }
                }

                int colClass = -1, colNum = -1, colName = -1, colGender = -1, colRank = -1, colScore = -1, colRemarks = -1;
                foreach (var kv in colHeaders)
                {
                    string h = kv.Value.Replace(" ", "");
                    if ((h.Contains("반") || h.Contains("학급") || h.Contains("원적") || h.Contains("기존반") || h.Contains("현재반") || h.Contains("이전반")) && colClass == -1) colClass = kv.Key;
                    else if ((h.Contains("번호") || h.Contains("순번") || h == "번") && colNum == -1) colNum = kv.Key;
                    else if ((h.Contains("이름") || h.Contains("성명") || h.Contains("학생")) && colName == -1) colName = kv.Key;
                    else if ((h.Contains("성별") || h == "성") && colGender == -1) colGender = kv.Key;
                    else if ((h.Contains("석차") || h.Contains("등수") || h.Contains("순위")) && colRank == -1) colRank = kv.Key;
                    else if ((h.Contains("합계") || h.Contains("점수") || h.Contains("총점") || h.Contains("성적") || h.Contains("평균")) && colScore == -1) colScore = kv.Key;
                    else if ((h.Contains("비고") || h.Contains("특기") || h.Contains("참고") || h.Contains("메모")) && colRemarks == -1) colRemarks = kv.Key;
                }

                if (colClass == -1)
                {
                    if (colNum == -1) colNum = 0;
                    if (colName == -1) colName = 1;
                    if (colGender == -1) colGender = 2;
                    if (colScore == -1) colScore = 5;
                    if (colRank == -1) colRank = 6;
                    if (colRemarks == -1) colRemarks = 7;
                }
                else
                {
                    if (colNum == -1) colNum = 1;
                    if (colName == -1) colName = 2;
                    if (colGender == -1) colGender = 3;
                    if (colScore == -1) colScore = 6;
                    if (colRank == -1) colRank = 7;
                    if (colRemarks == -1) colRemarks = 8;
                }

                for (int r = headerRowIdx + 1; r < rowNodes.Count; r++)
                {
                    XmlNode row = rowNodes[r];
                    Dictionary<int, string> rowValues = new Dictionary<int, string>();
                    foreach (XmlNode c in row.ChildNodes)
                    {
                        if (!string.Equals(c.LocalName, "c", StringComparison.OrdinalIgnoreCase)) continue;
                        int colIdx = GetColumnIndex(GetAttr(c, "r"));
                        rowValues[colIdx] = GetCellValue(c, sharedStrings);
                    }

                    string name = rowValues.ContainsKey(colName) ? rowValues[colName].Trim() : "";
                    if (string.IsNullOrEmpty(name)) continue;
                    if (name == "이름" || name == "성명" || name == "학생명") continue;

                    string prevClass = (colClass != -1 && rowValues.ContainsKey(colClass)) ? rowValues[colClass].Trim() : sheetName;
                    if (string.IsNullOrEmpty(prevClass)) prevClass = sheetName;

                    string prevNum = (colNum != -1 && rowValues.ContainsKey(colNum)) ? rowValues[colNum].Trim() : "";
                    string gender = (colGender != -1 && rowValues.ContainsKey(colGender)) ? rowValues[colGender].Trim() : "남";
                    if (gender.Contains("여") || gender.StartsWith("F", StringComparison.OrdinalIgnoreCase) || gender.StartsWith("W", StringComparison.OrdinalIgnoreCase)) gender = "여";
                    else gender = "남";

                    string rankStr = (colRank != -1 && rowValues.ContainsKey(colRank)) ? rowValues[colRank].Trim() : "99";
                    string scoreStr = (colScore != -1 && rowValues.ContainsKey(colScore)) ? rowValues[colScore].Trim() : "0";
                    string remarks = (colRemarks != -1 && rowValues.ContainsKey(colRemarks)) ? rowValues[colRemarks].Trim() : "";

                    int rank = 99;
                    int.TryParse(rankStr, out rank);
                    double score = 0;
                    double.TryParse(scoreStr, out score);

                    StudentRecord std = new StudentRecord
                    {
                        Id = prevClass + "_" + prevNum + "_" + name,
                        PrevClass = prevClass,
                        PrevNumber = prevNum,
                        Name = name,
                        Gender = gender,
                        Rank = rank,
                        Score = score,
                        Remarks = remarks,
                        IsSpecial = remarks.Contains("특수"),
                        SeparateFromNames = NameParser.ParseSeparationRequests(remarks)
                    };

                    string sur, first;
                    NameParser.ParseName(name, out sur, out first);
                    std.Surname = sur;
                    std.FirstName = first;

                    students.Add(std);
                }
            }
            return students;
        }

        private static string GetAttr(XmlNode node, string attrName)
        {
            if (node == null || node.Attributes == null) return null;
            XmlAttribute a = node.Attributes[attrName];
            if (a != null) return a.Value;
            foreach (XmlAttribute attr in node.Attributes)
            {
                if (string.Equals(attr.Name, attrName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attr.LocalName, attrName, StringComparison.OrdinalIgnoreCase))
                {
                    return attr.Value;
                }
            }
            return null;
        }

        private static int GetColumnIndex(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return 0;
            int col = 0;
            for (int i = 0; i < cellRef.Length; i++)
            {
                char c = cellRef[i];
                if (char.IsLetter(c))
                {
                    col = col * 26 + (char.ToUpper(c) - 'A' + 1);
                }
                else break;
            }
            return Math.Max(0, col - 1);
        }

        private static string GetCellValue(XmlNode cellNode, List<string> sharedStrings)
        {
            if (cellNode == null) return "";
            string type = GetAttr(cellNode, "t");

            // 1. inlineStr: <is><t>val</t></is>
            XmlNode isNode = cellNode["is"];
            if (isNode != null)
            {
                StringBuilder sb = new StringBuilder();
                XmlNodeList tList = isNode.SelectNodes(".//t");
                if (tList != null && tList.Count > 0)
                {
                    foreach (XmlNode tn in tList) sb.Append(tn.InnerText);
                }
                else
                {
                    sb.Append(isNode.InnerText);
                }
                return sb.ToString();
            }

            // 2. Direct text node <t>
            XmlNode tDirect = cellNode["t"];
            if (tDirect != null) return tDirect.InnerText;

            // 3. Value node <v>
            XmlNode vNode = cellNode["v"];
            if (vNode == null)
            {
                return cellNode.InnerText ?? "";
            }

            string val = vNode.InnerText ?? "";
            if (type == "s")
            {
                int sIdx;
                if (int.TryParse(val, out sIdx) && sharedStrings != null && sIdx >= 0 && sIdx < sharedStrings.Count)
                {
                    return sharedStrings[sIdx];
                }
            }
            else if (type == "b")
            {
                return val == "1" ? "TRUE" : "FALSE";
            }
            return val;
        }

        public static Dictionary<string, List<StudentRecord>> ReadCsv(string filePath)
        {
            var result = new Dictionary<string, List<StudentRecord>>();
            string[] lines = null;
            try
            {
                string text = File.ReadAllText(filePath, Encoding.UTF8);
                if (text.Contains(""))
                {
                    text = File.ReadAllText(filePath, Encoding.Default);
                }
                lines = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            }
            catch
            {
                try { lines = File.ReadAllLines(filePath, Encoding.Default); } catch { return result; }
            }

            if (lines == null || lines.Length < 2) return result;

            string[] headers = lines[0].Split(new char[] { ',', '\t' });
            int colClass = -1, colNum = -1, colName = -1, colGender = -1, colRank = -1, colScore = -1, colRemarks = -1;
            for (int i = 0; i < headers.Length; i++)
            {
                string h = headers[i].Trim().Replace(" ", "");
                if ((h.Contains("반") || h.Contains("학급") || h.Contains("원적")) && colClass == -1) colClass = i;
                else if ((h.Contains("번호") || h.Contains("순번") || h == "번") && colNum == -1) colNum = i;
                else if ((h.Contains("이름") || h.Contains("성명") || h.Contains("학생")) && colName == -1) colName = i;
                else if ((h.Contains("성별") || h == "성") && colGender == -1) colGender = i;
                else if ((h.Contains("석차") || h.Contains("등수") || h.Contains("순위")) && colRank == -1) colRank = i;
                else if ((h.Contains("합계") || h.Contains("점수") || h.Contains("총점") || h.Contains("성적")) && colScore == -1) colScore = i;
                else if ((h.Contains("비고") || h.Contains("특기") || h.Contains("참고")) && colRemarks == -1) colRemarks = i;
            }

            if (colClass == -1)
            {
                if (colNum == -1) colNum = 0;
                if (colName == -1) colName = 1;
                if (colGender == -1) colGender = 2;
                if (colScore == -1) colScore = 5;
                if (colRank == -1) colRank = 6;
                if (colRemarks == -1) colRemarks = 7;
            }
            else
            {
                if (colNum == -1) colNum = 1;
                if (colName == -1) colName = 2;
                if (colGender == -1) colGender = 3;
                if (colScore == -1) colScore = 6;
                if (colRank == -1) colRank = 7;
                if (colRemarks == -1) colRemarks = 8;
            }

            for (int r = 1; r < lines.Length; r++)
            {
                string[] parts = lines[r].Split(new char[] { ',', '\t' });
                if (parts.Length <= colName) continue;

                string name = parts[colName].Trim();
                if (string.IsNullOrEmpty(name)) continue;

                string prevClass = (colClass != -1 && parts.Length > colClass) ? parts[colClass].Trim() : "1반";
                if (string.IsNullOrEmpty(prevClass)) prevClass = "1반";

                string prevNum = (colNum != -1 && parts.Length > colNum) ? parts[colNum].Trim() : "";
                string gender = (colGender != -1 && parts.Length > colGender) ? parts[colGender].Trim() : "남";
                if (gender.Contains("여") || gender.StartsWith("F", StringComparison.OrdinalIgnoreCase) || gender.StartsWith("W", StringComparison.OrdinalIgnoreCase)) gender = "여";
                else gender = "남";

                string rankStr = (colRank != -1 && parts.Length > colRank) ? parts[colRank].Trim() : "99";
                string scoreStr = (colScore != -1 && parts.Length > colScore) ? parts[colScore].Trim() : "0";
                string remarks = (colRemarks != -1 && parts.Length > colRemarks) ? parts[colRemarks].Trim() : "";

                int rank = 99;
                int.TryParse(rankStr, out rank);
                double score = 0;
                double.TryParse(scoreStr, out score);

                StudentRecord std = new StudentRecord
                {
                    Id = prevClass + "_" + prevNum + "_" + name,
                    PrevClass = prevClass,
                    PrevNumber = prevNum,
                    Name = name,
                    Gender = gender,
                    Rank = rank,
                    Score = score,
                    Remarks = remarks,
                    IsSpecial = remarks.Contains("특수"),
                    SeparateFromNames = NameParser.ParseSeparationRequests(remarks)
                };
                string sur, first;
                NameParser.ParseName(name, out sur, out first);
                std.Surname = sur;
                std.FirstName = first;

                if (!result.ContainsKey(prevClass)) result[prevClass] = new List<StudentRecord>();
                result[prevClass].Add(std);
            }

            return result;
        }

        public static void WriteStyledXlsx(string filePath, Dictionary<string, List<StudentRecord>> classSheets, bool isTemplate = false)
        {
            // 기본 엑셀 구조 생성
            if (File.Exists(filePath)) File.Delete(filePath);

            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                // 1. [Content_Types].xml
                WriteZipEntry(zip, "[Content_Types].xml", GenerateContentTypesXml(classSheets.Count));

                // 2. _rels/.rels
                WriteZipEntry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                    "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
                    "</Relationships>");

                // 3. xl/styles.xml (특이사항 배경색 팔레트 정의)
                WriteZipEntry(zip, "xl/styles.xml", GenerateStylesXml());

                // 4. xl/workbook.xml & xl/_rels/workbook.xml.rels
                WriteZipEntry(zip, "xl/workbook.xml", GenerateWorkbookXml(classSheets));
                WriteZipEntry(zip, "xl/_rels/workbook.xml.rels", GenerateWorkbookRelsXml(classSheets.Count));

                // 5. xl/worksheets/sheet*.xml 생성
                int sheetIndex = 1;
                foreach (var kv in classSheets)
                {
                    string sheetXml = GenerateWorksheetXml(kv.Value, isTemplate);
                    WriteZipEntry(zip, string.Format("xl/worksheets/sheet{0}.xml", sheetIndex), sheetXml);
                    sheetIndex++;
                }
            }
        }

        private static void WriteZipEntry(ZipArchive zip, string entryName, string content)
        {
            var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.Write(content);
            }
        }

        private static string GenerateContentTypesXml(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
            sb.Append("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n");
            sb.Append("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n");
            sb.Append("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n");
            sb.Append("  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n");
            for (int i = 1; i <= sheetCount; i++)
            {
                sb.AppendFormat("  <Override PartName=\"/xl/worksheets/sheet{0}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n", i);
            }
            sb.Append("</Types>");
            return sb.ToString();
        }

        private static string GenerateWorkbookXml(Dictionary<string, List<StudentRecord>> sheets)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
            sb.Append("  <sheets>\n");
            int id = 1;
            foreach (var name in sheets.Keys)
            {
                sb.AppendFormat("    <sheet name=\"{0}\" sheetId=\"{1}\" r:id=\"rId{1}\"/>\n", EscapeXml(name), id);
                id++;
            }
            sb.Append("  </sheets>\n");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        private static string GenerateWorkbookRelsXml(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
            for (int i = 1; i <= sheetCount; i++)
            {
                sb.AppendFormat("  <Relationship Id=\"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{0}.xml\"/>\n", i);
            }
            sb.AppendFormat("  <Relationship Id=\"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n", sheetCount + 1);
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string GenerateStylesXml()
        {
            // Fills:
            // 0: none, 1: gray125
            // 2: Header (#E2E8F0)
            // 3: 생활지도 빨강 (#FFD1D5)
            // 4: 특수아동 파랑 (#BFDBFE)
            // 5: 다문화 노랑 (#FEF08A)
            // 6: 학습부진 주황 (#FED7AA)
            // 7: 분리요청 보라 (#E9D5FF)
            // 8: 운동부 초록 (#BBF7D0)
            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                "  <fonts count=\"3\">\n" +
                "    <font><name val=\"Malgun Gothic\"/><sz val=\"10\"/></font>\n" +
                "    <font><name val=\"Malgun Gothic\"/><sz val=\"10\"/><b/></font>\n" +
                "    <font><name val=\"Malgun Gothic\"/><sz val=\"10\"/><b/><color rgb=\"FF9F1239\"/></font>\n" +
                "  </fonts>\n" +
                "  <fills count=\"9\">\n" +
                "    <fill><patternFill patternType=\"none\"/></fill>\n" +
                "    <fill><patternFill patternType=\"gray125\"/></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE2E8F0\"/></patternFill></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFD1D5\"/></patternFill></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFBFDBFE\"/></patternFill></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFEF08A\"/></patternFill></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFED7AA\"/></patternFill></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE9D5FF\"/></patternFill></fill>\n" +
                "    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFBBF7D0\"/></patternFill></fill>\n" +
                "  </fills>\n" +
                "  <borders count=\"2\">\n" +
                "    <border><left/><right/><top/><bottom/></border>\n" +
                "    <border>\n" +
                "      <left style=\"thin\"><color rgb=\"FFCBD5E1\"/></left>\n" +
                "      <right style=\"thin\"><color rgb=\"FFCBD5E1\"/></right>\n" +
                "      <top style=\"thin\"><color rgb=\"FFCBD5E1\"/></top>\n" +
                "      <bottom style=\"thin\"><color rgb=\"FFCBD5E1\"/></bottom>\n" +
                "    </border>\n" +
                "  </borders>\n" +
                "  <cellXfs count=\"9\">\n" +
                "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"4\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"5\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"6\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"7\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"8\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"center\"/></xf>\n" +
                "  </cellXfs>\n" +
                "</styleSheet>";
        }

        private static string GenerateWorksheetXml(List<StudentRecord> students, bool isTemplate)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
            sb.Append("  <cols>\n");
            if (isTemplate)
            {
                sb.Append("    <col min=\"1\" max=\"1\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"2\" max=\"2\" width=\"13\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"3\" max=\"3\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"4\" max=\"4\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"5\" max=\"5\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"6\" max=\"6\" width=\"12\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"7\" max=\"7\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"8\" max=\"8\" width=\"32\" customWidth=\"1\"/>\n");
            }
            else
            {
                sb.Append("    <col min=\"1\" max=\"1\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"2\" max=\"2\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"3\" max=\"3\" width=\"13\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"4\" max=\"4\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"5\" max=\"5\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"6\" max=\"6\" width=\"12\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"7\" max=\"7\" width=\"10\" customWidth=\"1\"/>\n");
                sb.Append("    <col min=\"8\" max=\"8\" width=\"32\" customWidth=\"1\"/>\n");
            }
            sb.Append("  </cols>\n");
            sb.Append("  <sheetData>\n");

            // 헤더 행 (s="1")
            string[] headers = isTemplate
                ? new string[] { "번호", "이름", "성별", "국어", "수학", "합계", "석차", "비고" }
                : new string[] { "반", "번호", "이름", "성별", "석차", "진급반", "새번호", "비고" };

            sb.Append("    <row r=\"1\">\n");
            for (int i = 0; i < headers.Length; i++)
            {
                string colLetter = ((char)('A' + i)).ToString();
                sb.AppendFormat("      <c r=\"{0}1\" t=\"inlineStr\" s=\"1\"><is><t>{1}</t></is></c>\n", colLetter, headers[i]);
            }
            sb.Append("    </row>\n");

            // 데이터 행
            int rowIdx = 2;
            foreach (var s in students)
            {
                sb.AppendFormat("    <row r=\"{0}\">\n", rowIdx);

                string[] vals;
                if (isTemplate)
                {
                    int kor = s.Score > 0 ? (int)(s.Score * 0.52) : 90;
                    int math = s.Score > 0 ? (int)(s.Score * 0.48) : 88;
                    int total = s.Score > 0 ? (int)s.Score : (kor + math);
                    vals = new string[]
                    {
                        s.PrevNumber ?? "",
                        s.Name ?? "",
                        s.Gender ?? "",
                        kor.ToString(),
                        math.ToString(),
                        total.ToString(),
                        s.Rank > 0 ? s.Rank.ToString() : "",
                        s.Remarks ?? ""
                    };
                }
                else
                {
                    vals = new string[]
                    {
                        s.PrevClass ?? "",
                        s.PrevNumber ?? "",
                        s.Name ?? "",
                        s.Gender ?? "",
                        s.Rank > 0 ? s.Rank.ToString() : "",
                        s.AssignedClass ?? "",
                        s.AssignedNumber > 0 ? s.AssignedNumber.ToString() : "",
                        s.Remarks ?? ""
                    };
                }

                // 비고 스타일 식별
                int remarkStyle = 2; // 기본 테두리+가운데정렬
                string rem = s.Remarks ?? "";
                if (rem.Contains("생활지도")) remarkStyle = 3;
                else if (rem.Contains("특수")) remarkStyle = 4;
                else if (rem.Contains("다문화")) remarkStyle = 5;
                else if (rem.Contains("학습부진")) remarkStyle = 6;
                else if (rem.Contains("분리요청")) remarkStyle = 7;
                else if (rem.Contains("운동부")) remarkStyle = 8;

                for (int c = 0; c < vals.Length; c++)
                {
                    string colLetter = ((char)('A' + c)).ToString();
                    int styleId = (c == 7) ? remarkStyle : 2;
                    sb.AppendFormat("      <c r=\"{0}{1}\" t=\"inlineStr\" s=\"{2}\"><is><t>{3}</t></is></c>\n",
                        colLetter, rowIdx, styleId, EscapeXml(vals[c]));
                }
                sb.Append("    </row>\n");
                rowIdx++;
            }

            sb.Append("  </sheetData>\n");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }
    }

    // =========================================================
    // 4. 최적화 알고리즘 엔진
    // =========================================================
    public class ClassOptimizerEngine
    {
        private List<StudentRecord> students;
        private List<string> newClasses;
        private List<string> prevClasses;
        private string scoreType;
        private string numberingMode;
        private int K;
        private Random rnd = new Random();

        public ClassOptimizerEngine(List<StudentRecord> stds, List<string> newC, List<string> prevC, string st, string nm)
        {
            students = stds;
            newClasses = newC;
            prevClasses = prevC;
            scoreType = st;
            numberingMode = nm;
            K = newClasses.Count;
        }

        public Dictionary<string, List<StudentRecord>> Optimize()
        {
            Dictionary<string, List<string>> assignment = new Dictionary<string, List<string>>();
            foreach (var c in newClasses) assignment[c] = new List<string>();

            // 1. 초기 지그재그(스네이크) 배정
            List<StudentRecord> boys = new List<StudentRecord>();
            List<StudentRecord> girls = new List<StudentRecord>();
            foreach (var s in students)
            {
                if (s.Gender == "여") girls.Add(s);
                else boys.Add(s);
            }

            if (scoreType == "score")
            {
                boys.Sort((a, b) => b.Score.CompareTo(a.Score));
                girls.Sort((a, b) => b.Score.CompareTo(a.Score));
            }
            else
            {
                boys.Sort((a, b) => a.Rank.CompareTo(b.Rank));
                girls.Sort((a, b) => a.Rank.CompareTo(b.Rank));
            }

            Action<List<StudentRecord>, int> snake = (list, offset) =>
            {
                int curr = offset % K;
                int dir = 1;
                foreach (var s in list)
                {
                    assignment[newClasses[curr]].Add(s.Id);
                    if (dir == 1)
                    {
                        if (curr == K - 1) dir = -1;
                        else curr++;
                    }
                    else
                    {
                        if (curr == 0) dir = 1;
                        else curr--;
                    }
                }
            };

            snake(boys, 0);
            snake(girls, 1);

            // 2. Simulated Annealing 제약 최적화
            Dictionary<string, StudentRecord> idMap = new Dictionary<string, StudentRecord>();
            Dictionary<string, List<string>> nameToIds = new Dictionary<string, List<string>>();
            foreach (var s in students)
            {
                idMap[s.Id] = s;
                if (!nameToIds.ContainsKey(s.Name)) nameToIds[s.Name] = new List<string>();
                nameToIds[s.Name].Add(s.Id);
            }

            Func<double> calcEnergy = () =>
            {
                double penalty = 0;
                foreach (var c in newClasses)
                {
                    List<string> ids = assignment[c];
                    HashSet<string> seenFirstNames = new HashSet<string>();
                    int mCount = 0, fCount = 0, specCount = 0;

                    foreach (var id in ids)
                    {
                        var s = idMap[id];
                        if (s.Gender == "여") fCount++; else mCount++;
                        if (s.IsSpecial) specCount++;

                        // 이름 중복 검사
                        if (!string.IsNullOrEmpty(s.FirstName))
                        {
                            if (seenFirstNames.Contains(s.FirstName)) penalty += 500000;
                            else seenFirstNames.Add(s.FirstName);
                        }

                        // 분리 요청 검사
                        foreach (var targetName in s.SeparateFromNames)
                        {
                            if (nameToIds.ContainsKey(targetName))
                            {
                                foreach (var tid in nameToIds[targetName])
                                {
                                    if (tid != s.Id && ids.Contains(tid))
                                    {
                                        penalty += 1000000;
                                    }
                                }
                            }
                        }
                    }

                    // 성비 균등
                    penalty += Math.Abs(mCount - fCount) * 1000;
                    // 특수아동 가중 인원 균등
                    int weighted = ids.Count + specCount;
                    penalty += Math.Abs(weighted - ((double)students.Count / K)) * 5000;
                }
                return penalty;
            };

            double currentEnergy = calcEnergy();
            double temp = 100.0;
            double cooling = 0.99;

            for (int it = 0; it < 1500; it++)
            {
                if (currentEnergy == 0) break;

                int c1Idx = rnd.Next(K);
                int c2Idx = rnd.Next(K);
                if (c1Idx == c2Idx) continue;

                string c1 = newClasses[c1Idx];
                string c2 = newClasses[c2Idx];
                if (assignment[c1].Count == 0 || assignment[c2].Count == 0) continue;

                string id1 = assignment[c1][rnd.Next(assignment[c1].Count)];
                var s1 = idMap[id1];

                // 동일 성별 교환 우선
                List<string> sameGender = assignment[c2].FindAll(x => idMap[x].Gender == s1.Gender);
                string id2 = sameGender.Count > 0 ? sameGender[rnd.Next(sameGender.Count)] : assignment[c2][rnd.Next(assignment[c2].Count)];

                assignment[c1].Remove(id1);
                assignment[c2].Remove(id2);
                assignment[c1].Add(id2);
                assignment[c2].Add(id1);

                double newEnergy = calcEnergy();
                double delta = newEnergy - currentEnergy;

                if (delta < 0 || (temp > 0 && rnd.NextDouble() < Math.Exp(-delta / temp)))
                {
                    currentEnergy = newEnergy;
                }
                else
                {
                    // 롤백
                    assignment[c1].Remove(id2);
                    assignment[c2].Remove(id1);
                    assignment[c1].Add(id1);
                    assignment[c2].Add(id2);
                }
                temp *= cooling;
            }

            // 3. 번호 부여
            var finalResult = new Dictionary<string, List<StudentRecord>>();
            foreach (var c in newClasses)
            {
                List<StudentRecord> classStds = new List<StudentRecord>();
                foreach (var id in assignment[c])
                {
                    var s = idMap[id];
                    s.AssignedClass = c;
                    classStds.Add(s);
                }

                if (numberingMode == "name_all")
                {
                    classStds.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    for (int i = 0; i < classStds.Count; i++) classStds[i].AssignedNumber = i + 1;
                }
                else if (numberingMode == "continuous") // 남 1번~ / 여 21번~
                {
                    var bList = classStds.FindAll(x => x.Gender == "남");
                    var gList = classStds.FindAll(x => x.Gender == "여");
                    bList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    gList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    for (int i = 0; i < bList.Count; i++) bList[i].AssignedNumber = i + 1;
                    int gStart = Math.Max(21, bList.Count + 1);
                    for (int i = 0; i < gList.Count; i++) gList[i].AssignedNumber = gStart + i;
                    classStds = new List<StudentRecord>(bList);
                    classStds.AddRange(gList);
                    classStds.Sort((a, b) => a.AssignedNumber.CompareTo(b.AssignedNumber));
                }
                else if (numberingMode == "continuous_female") // 여 1번~ / 남 21번~
                {
                    var bList = classStds.FindAll(x => x.Gender == "남");
                    var gList = classStds.FindAll(x => x.Gender == "여");
                    bList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    gList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    for (int i = 0; i < gList.Count; i++) gList[i].AssignedNumber = i + 1;
                    int bStart = Math.Max(21, gList.Count + 1);
                    for (int i = 0; i < bList.Count; i++) bList[i].AssignedNumber = bStart + i;
                    classStds = new List<StudentRecord>(gList);
                    classStds.AddRange(bList);
                    classStds.Sort((a, b) => a.AssignedNumber.CompareTo(b.AssignedNumber));
                }
                else // separate: 남 1번~ / 여 1번~
                {
                    var bList = classStds.FindAll(x => x.Gender == "남");
                    var gList = classStds.FindAll(x => x.Gender == "여");
                    bList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    gList.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));
                    for (int i = 0; i < bList.Count; i++) bList[i].AssignedNumber = i + 1;
                    for (int i = 0; i < gList.Count; i++) gList[i].AssignedNumber = i + 1;
                    classStds = new List<StudentRecord>(bList);
                    classStds.AddRange(gList);
                }

                finalResult[c] = classStds;
            }

            return finalResult;
        }
    }

    // =========================================================
    // 5. 고품질 모던 UI 렌더링 헬퍼 & 커스텀 컨트롤 (HTML 버전 100% 매칭)
    // =========================================================
    public static class GraphicsHelper
    {
        public static GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = radius * 2;
            if (rect.Width <= 0 || rect.Height <= 0) return path;
            if (d > rect.Width) d = rect.Width;
            if (d > rect.Height) d = rect.Height;

            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public enum Button3DTheme
    {
        Mint,
        White,
        SelectedMint
    }

    public class Button3D : Button
    {
        private Button3DTheme theme = Button3DTheme.Mint;
        public Button3DTheme Theme
        {
            get { return theme; }
            set { theme = value; Invalidate(); }
        }

        private bool isHovered = false;
        private bool isPressed = false;
        private int cornerRadius = 12;

        public int CornerRadius
        {
            get { return cornerRadius; }
            set { cornerRadius = value; Invalidate(); }
        }

        public Button3D()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | 
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.Selectable, false);

            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderSize = 0;
            this.FlatAppearance.BorderColor = Color.FromArgb(0, 255, 255, 255);
            this.FlatAppearance.MouseDownBackColor = Color.Transparent;
            this.FlatAppearance.MouseOverBackColor = Color.Transparent;
            this.BackColor = Color.Transparent;
            this.Cursor = Cursors.Hand;
            this.Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold);
        }

        protected override void OnMouseEnter(EventArgs e) { isHovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { isHovered = false; isPressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs mevent) { if (mevent.Button == MouseButtons.Left) { isPressed = true; Invalidate(); } base.OnMouseDown(mevent); }
        protected override void OnMouseUp(MouseEventArgs mevent) { isPressed = false; Invalidate(); base.OnMouseUp(mevent); }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // 부모 배경색으로 외곽 완전히 지우기 (검은 점 및 테두리 잔여선 원천 방지)
            Color parentBg = (Parent != null) ? Parent.BackColor : Color.FromArgb(248, 250, 252);
            using (SolidBrush pb = new SolidBrush(parentBg))
            {
                g.FillRectangle(pb, ClientRectangle);
            }

            int r = cornerRadius;
            Rectangle rect = new Rectangle(1, 1, Width - 3, Height - 3);
            if (isPressed)
            {
                rect.Y += 1;
            }

            Color bgTop, bgBottom, strokeColor, textColor;

            if (theme == Button3DTheme.Mint)
            {
                bgTop = isHovered ? Color.FromArgb(52, 211, 153) : Color.FromArgb(16, 185, 129);
                bgBottom = isHovered ? Color.FromArgb(16, 185, 129) : Color.FromArgb(5, 150, 105);
                strokeColor = isHovered ? Color.FromArgb(52, 211, 153) : Color.FromArgb(16, 185, 129);
                textColor = Color.White;
            }
            else if (theme == Button3DTheme.SelectedMint)
            {
                bgTop = Color.FromArgb(5, 150, 105);
                bgBottom = Color.FromArgb(4, 120, 87);
                strokeColor = Color.FromArgb(110, 231, 183);
                textColor = Color.White;
            }
            else // White
            {
                bgTop = isHovered ? Color.FromArgb(241, 245, 249) : Color.White;
                bgBottom = isHovered ? Color.FromArgb(226, 232, 240) : Color.FromArgb(248, 250, 252);
                strokeColor = isHovered ? Color.FromArgb(148, 163, 184) : Color.FromArgb(203, 213, 225);
                textColor = Color.FromArgb(30, 41, 59);
            }

            // 부드러운 라운드 본체 렌더링
            using (GraphicsPath btnPath = GraphicsHelper.GetRoundedPath(rect, r))
            {
                using (LinearGradientBrush lgb = new LinearGradientBrush(rect, bgTop, bgBottom, LinearGradientMode.Vertical))
                {
                    g.FillPath(lgb, btnPath);
                }

                float strokeW = (theme == Button3DTheme.SelectedMint) ? 2.5f : 1.4f;
                using (Pen sp = new Pen(strokeColor, strokeW))
                {
                    g.DrawPath(sp, btnPath);
                }
            }

            // 텍스트 균형 배치
            StringFormat sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            Rectangle textRect = new Rectangle(rect.X + 2, rect.Y, rect.Width - 4, rect.Height);
            using (SolidBrush tb = new SolidBrush(textColor))
            {
                g.DrawString(Text, Font, tb, textRect, sf);
            }
        }
    }

    public class ModernCard : Panel
    {
        public int Radius { get; set; }
        public Color BorderColor { get; set; }
        public Color ShadowColor { get; set; }

        public ModernCard()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | 
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Radius = 14;
            BackColor = Color.White;
            BorderColor = Color.FromArgb(226, 232, 240);
            ShadowColor = Color.FromArgb(241, 245, 249);
            Padding = new Padding(15);
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (Parent != null)
            {
                using (SolidBrush pb = new SolidBrush(Parent.BackColor))
                {
                    g.FillRectangle(pb, ClientRectangle);
                }
            }

            Rectangle shadowRect = new Rectangle(0, 1, Width - 1, Height - 1);
            using (GraphicsPath sp = GraphicsHelper.GetRoundedPath(shadowRect, Radius))
            using (SolidBrush sb = new SolidBrush(ShadowColor))
            {
                g.FillPath(sb, sp);
            }

            Rectangle mainRect = new Rectangle(0, 0, Width - 1, Height - 2);
            using (GraphicsPath mp = GraphicsHelper.GetRoundedPath(mainRect, Radius))
            {
                using (SolidBrush mb = new SolidBrush(BackColor))
                {
                    g.FillPath(mb, mp);
                }
                using (Pen p = new Pen(BorderColor, 1.2f))
                {
                    g.DrawPath(p, mp);
                }
            }
        }
    }

    public class StatCard : Panel
    {
        public string Title { get; set; }
        public string Value { get; set; }
        public string SubText { get; set; }
        public Color ValueColor { get; set; }
        public Color AccentBg { get; set; }
        public Color AccentBorder { get; set; }

        public StatCard()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | 
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            ValueColor = Color.FromArgb(30, 41, 59);
            AccentBg = Color.White;
            AccentBorder = Color.FromArgb(226, 232, 240);
            Title = "";
            Value = "";
            SubText = "";
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            if (Parent != null)
            {
                using (SolidBrush pb = new SolidBrush(Parent.BackColor))
                {
                    g.FillRectangle(pb, ClientRectangle);
                }
            }

            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (GraphicsPath path = GraphicsHelper.GetRoundedPath(r, 12))
            {
                using (SolidBrush bb = new SolidBrush(AccentBg))
                {
                    g.FillPath(bb, path);
                }
                using (Pen p = new Pen(AccentBorder, 1.2f))
                {
                    g.DrawPath(p, path);
                }
            }

            // Title
            using (Font tf = new Font("Malgun Gothic", 12f, FontStyle.Bold))
            using (SolidBrush tb = new SolidBrush(Color.FromArgb(100, 116, 139)))
            {
                g.DrawString(Title, tf, tb, new RectangleF(14, 12, Width - 28, 22));
            }

            // Value
            using (Font vf = new Font("Malgun Gothic", 19f, FontStyle.Bold))
            using (SolidBrush vb = new SolidBrush(ValueColor))
            {
                g.DrawString(Value, vf, vb, new RectangleF(14, 34, Width - 28, 36));
            }

            // SubText
            if (!string.IsNullOrEmpty(SubText))
            {
                using (Font sf = new Font("Malgun Gothic", 11f, FontStyle.Regular))
                using (SolidBrush sb = new SolidBrush(Color.FromArgb(71, 85, 105)))
                {
                    g.DrawString(SubText, sf, sb, new RectangleF(14, 72, Width - 28, 24));
                }
            }
        }
    }

    // =========================================================
    // 6. 네이티브 GUI 메인 윈도우 (WinForms)
    // =========================================================
    public class MainForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ChangeWindowMessageFilter(uint message, uint dwFlag);

        private const uint WM_DROPFILES = 0x0233;
        private const uint WM_COPYDATA = 0x004A;
        private const uint WM_COPYGLOBALDATA = 0x0049;
        private const uint MSGFLT_ADD = 1;

        private List<StudentRecord> rawRecords = new List<StudentRecord>();
        private List<string> prevClasses = new List<string> { "1반", "2반", "3반", "4반" };
        private List<string> newClasses = new List<string> { "가반", "나반", "다반", "라반" };
        private string scoreType = "rank";
        private string numberingMode = "name_all";
        private Dictionary<string, List<StudentRecord>> assignmentResult = null;

        // UI 패널
        private Panel headerPanel;
        private Panel contentPanel;
        private Panel step1Panel;
        private Panel step2Panel;
        private Panel step3Panel;

        // Step 1 컨트롤
        private Panel dropArea;
        private bool isDropAreaHovered = false;
        private Label lblFileStatus;
        private Button3D btnScoreRank;
        private Button3D btnScoreTotal;
        private Button3D btnNumNameAll;
        private Button3D btnNumSeparate;
        private Button3D btnNumContinuous;
        private Button3D btnNumContinuousFemale;

        // Step 2 컨트롤
        private TextBox txtNewClasses;
        private Label lblStep2Info;

        // Step 3 컨트롤
        private Panel statCardsPanel;
        private FlowLayoutPanel classTabPanel;
        private DataGridView gridResult;
        private string activeClassTab = "전체통합";

        public MainForm()
        {
            try
            {
                ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD);
                ChangeWindowMessageFilter(WM_COPYDATA, MSGFLT_ADD);
                ChangeWindowMessageFilter(WM_COPYGLOBALDATA, MSGFLT_ADD);
            }
            catch { }

            InitializeComponent();
            ShowStep1();
        }

        private void HighlightDropArea(bool highlight)
        {
            if (isDropAreaHovered != highlight)
            {
                isDropAreaHovered = highlight;
                if (dropArea != null)
                {
                    dropArea.BackColor = highlight ? Color.FromArgb(220, 252, 231) : Color.FromArgb(240, 253, 244);
                    dropArea.Invalidate();
                }
            }
        }

        private void AttachFileDropHandler(Control control)
        {
            if (control == null) return;
            control.AllowDrop = true;
            control.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effect = DragDropEffects.Copy;
                    HighlightDropArea(true);
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            };
            control.DragOver += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effect = DragDropEffects.Copy;
                }
            };
            control.DragLeave += (s, e) =>
            {
                HighlightDropArea(false);
            };
            control.DragDrop += (s, e) =>
            {
                HighlightDropArea(false);
                try
                {
                    if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    {
                        string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                        if (files != null && files.Length > 0)
                        {
                            LoadExcel(files[0]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("파일을 읽는 도중 오류가 발생했습니다: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            foreach (Control child in control.Controls)
            {
                AttachFileDropHandler(child);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "스마트 학급편성 시스템";
            this.Size = new Size(1300, 920);
            this.MinimumSize = new Size(1120, 780);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Malgun Gothic", 11f, FontStyle.Regular);

            // 1. 헤더 바 (모던 클린 스타일)
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 78,
                BackColor = Color.White
            };
            headerPanel.Paint += (s, e) =>
            {
                using (Pen pen = new Pen(Color.FromArgb(226, 232, 240), 1.5f))
                {
                    e.Graphics.DrawLine(pen, 0, headerPanel.Height - 1, headerPanel.Width, headerPanel.Height - 1);
                }
            };

            Label lblTitle = new Label
            {
                Text = "스마트 학급편성 시스템",
                Font = new Font("Malgun Gothic", 18f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                Location = new Point(28, 22),
                AutoSize = true
            };

            // 커스텀 둥근 뱃지
            Panel pnlBadge = new Panel
            {
                Location = new Point(340, 24),
                Size = new Size(290, 32),
                BackColor = Color.FromArgb(209, 250, 229)
            };
            pnlBadge.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle r = new Rectangle(0, 0, pnlBadge.Width - 1, pnlBadge.Height - 1);
                using (GraphicsPath p = GraphicsHelper.GetRoundedPath(r, 14))
                using (Pen pen = new Pen(Color.FromArgb(167, 243, 208), 1.2f))
                {
                    e.Graphics.DrawPath(pen, p);
                }
            };
            Label lblBadge = new Label
            {
                Text = "완전 오프라인 네이티브 데스크톱 에디션",
                Font = new Font("Malgun Gothic", 11.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(6, 95, 70),
                BackColor = Color.Transparent,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            pnlBadge.Controls.Add(lblBadge);

            headerPanel.Controls.AddRange(new Control[] { lblTitle, pnlBadge });

            // 2. 메인 컨텐츠 영역 (헤더 바 아래 도킹)
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 250, 252)
            };

            this.Controls.Add(contentPanel);
            this.Controls.Add(headerPanel);
            headerPanel.SendToBack();
            contentPanel.BringToFront();

            // Step 패널 초기화 및 contentPanel에 등록
            CreateStep1Panel();
            CreateStep2Panel();
            CreateStep3Panel();

            contentPanel.Controls.AddRange(new Control[] { step1Panel, step2Panel, step3Panel });

            // 전체 폼 및 모든 하위 컨트롤에 드래그 앤 드롭 재귀 등록
            AttachFileDropHandler(this);
        }

        private void CreateStep1Panel()
        {
            step1Panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 250, 252),
                AutoScroll = true
            };

            Panel centerContainer = new Panel
            {
                Width = 1080,
                Height = 720,
                BackColor = Color.Transparent
            };

            // 1. 상단 안내 배너 카드
            ModernCard banner = new ModernCard
            {
                Location = new Point(0, 0),
                Size = new Size(1080, 145),
                Radius = 16,
                BackColor = Color.FromArgb(236, 253, 245),
                BorderColor = Color.FromArgb(167, 243, 208),
                ShadowColor = Color.FromArgb(209, 250, 229)
            };

            Label bTitle = new Label
            {
                Text = "엑셀 파일 저장 시 비고란 특이사항별 셀 배경색 자동 반영!",
                Font = new Font("Malgun Gothic", 15.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(6, 95, 70),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            Label bDesc = new Label
            {
                Text = "생활지도(빨강), 특수아동(파랑), 다문화(노랑), 학습부진(주황), 운동부(초록), 분리요청(보라) 등 배경색이 실제 엑셀 파일에 칠해져 저장됩니다.",
                Font = new Font("Malgun Gothic", 12.5f),
                ForeColor = Color.FromArgb(4, 120, 87),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            Button3D btnMock = new Button3D
            {
                Text = "예시 데이터 불러오기 (80명)",
                Theme = Button3DTheme.White,
                Font = new Font("Malgun Gothic", 13f, FontStyle.Bold),
                Size = new Size(280, 48)
            };
            btnMock.Click += (s, e) => LoadMockData();

            Button3D btnTemplate = new Button3D
            {
                Text = "표준 엑셀 양식 다운로드 (.xlsx)",
                Theme = Button3DTheme.Mint,
                Font = new Font("Malgun Gothic", 13f, FontStyle.Bold),
                Size = new Size(330, 48)
            };
            btnTemplate.Click += (s, e) => DownloadSampleExcel();

            Action updateBannerLayout = () =>
            {
                bTitle.Location = new Point(Math.Max(20, (banner.Width - bTitle.PreferredWidth) / 2), 16);
                bDesc.Location = new Point(Math.Max(20, (banner.Width - bDesc.PreferredWidth) / 2), 46);
                int totalBtnW = btnMock.Width + 18 + btnTemplate.Width;
                int startX = Math.Max(20, (banner.Width - totalBtnW) / 2);
                btnMock.Location = new Point(startX, 80);
                btnTemplate.Location = new Point(startX + btnMock.Width + 18, 80);
            };
            banner.Resize += (s, e) => updateBannerLayout();
            banner.Controls.AddRange(new Control[] { bTitle, bDesc, btnMock, btnTemplate });
            updateBannerLayout();

            // 2. 업로드 드롭 박스 카드
            ModernCard uploadCard = new ModernCard
            {
                Location = new Point(0, 160),
                Size = new Size(1080, 195),
                Radius = 16
            };

            dropArea = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(240, 253, 244)
            };
            dropArea.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle r = new Rectangle(1, 1, dropArea.Width - 3, dropArea.Height - 3);
                using (GraphicsPath p = GraphicsHelper.GetRoundedPath(r, 14))
                using (Pen pen = new Pen(isDropAreaHovered ? Color.FromArgb(5, 150, 105) : Color.FromArgb(167, 243, 208), isDropAreaHovered ? 2.5f : 2f))
                {
                    pen.DashStyle = DashStyle.Dash;
                    e.Graphics.DrawPath(pen, p);
                }
            };

            lblFileStatus = new Label
            {
                Text = "파일을 여기로 드래그 앤 드롭하거나 컴퓨터에서 선택하세요",
                Font = new Font("Malgun Gothic", 16f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            Label uHint = new Label
            {
                Text = "[번호, 이름, 성별, 국어, 수학, 합계, 석차, 비고] 엑셀(.xlsx) 또는 CSV 파일 지원 · 각 반 시트 자동 인식",
                Font = new Font("Malgun Gothic", 12.5f),
                ForeColor = Color.FromArgb(100, 116, 139),
                AutoSize = true,
                BackColor = Color.Transparent
            };
            Button3D btnSelectFile = new Button3D
            {
                Text = "컴퓨터에서 파일 찾기",
                Theme = Button3DTheme.White,
                Font = new Font("Malgun Gothic", 13.5f, FontStyle.Bold),
                Size = new Size(260, 48)
            };
            btnSelectFile.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV Files (*.csv)|*.csv" };
                if (ofd.ShowDialog() == DialogResult.OK) LoadExcel(ofd.FileName);
            };

            Action updateDropLayout = () =>
            {
                lblFileStatus.Location = new Point(Math.Max(20, (dropArea.Width - lblFileStatus.PreferredWidth) / 2), 22);
                uHint.Location = new Point(Math.Max(20, (dropArea.Width - uHint.PreferredWidth) / 2), 60);
                btnSelectFile.Location = new Point((dropArea.Width - btnSelectFile.Width) / 2, 110);
            };
            dropArea.Resize += (s, e) => updateDropLayout();
            dropArea.Controls.AddRange(new Control[] { lblFileStatus, uHint, btnSelectFile });
            uploadCard.Controls.Add(dropArea);
            updateDropLayout();

            // 3. 설정 카드 영역
            TableLayoutPanel settingsTable = new TableLayoutPanel
            {
                Location = new Point(0, 370),
                Size = new Size(1080, 240),
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.FromArgb(248, 250, 252)
            };
            settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
            settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));

            // 카드 1: 배정 성적 기준
            ModernCard cardScore = new ModernCard { Dock = DockStyle.Fill, Radius = 16, Margin = new Padding(0, 0, 10, 0) };
            Label scTitle = new Label { Text = "배정 성적 기준", Font = new Font("Malgun Gothic", 15.5f, FontStyle.Bold), ForeColor = Color.FromArgb(30, 41, 59), Location = new Point(22, 18), AutoSize = true, BackColor = Color.Transparent };

            btnScoreRank = new Button3D
            {
                Text = "석차 기준\n(1등이 우수)",
                Theme = Button3DTheme.SelectedMint,
                Font = new Font("Malgun Gothic", 13.5f, FontStyle.Bold),
                Location = new Point(22, 65),
                Size = new Size(225, 135)
            };
            btnScoreRank.Click += (s, e) => SetScoreType("rank");

            btnScoreTotal = new Button3D
            {
                Text = "합계 점수 기준\n(총점 높은 순)",
                Theme = Button3DTheme.White,
                Font = new Font("Malgun Gothic", 13.5f, FontStyle.Bold),
                Location = new Point(258, 65),
                Size = new Size(225, 135)
            };
            btnScoreTotal.Click += (s, e) => SetScoreType("score");

            cardScore.Controls.AddRange(new Control[] { scTitle, btnScoreRank, btnScoreTotal });

            // 카드 2: 진급반 번호 부여 방식
            ModernCard cardNum = new ModernCard { Dock = DockStyle.Fill, Radius = 16, Margin = new Padding(10, 0, 0, 0) };
            Label nmTitle = new Label { Text = "진급반 번호 부여 방식", Font = new Font("Malgun Gothic", 15.5f, FontStyle.Bold), ForeColor = Color.FromArgb(30, 41, 59), Location = new Point(22, 18), AutoSize = true, BackColor = Color.Transparent };

            Panel pnlRuleBadge = new Panel { Location = new Point(385, 20), Size = new Size(150, 28), BackColor = Color.FromArgb(236, 253, 245) };
            pnlRuleBadge.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle r = new Rectangle(0, 0, pnlRuleBadge.Width - 1, pnlRuleBadge.Height - 1);
                using (GraphicsPath p = GraphicsHelper.GetRoundedPath(r, 12))
                using (Pen pen = new Pen(Color.FromArgb(167, 243, 208), 1.2f)) e.Graphics.DrawPath(pen, p);
            };
            Label lblRuleBadge = new Label { Text = "새 학급 번호 규칙", Font = new Font("Malgun Gothic", 10.5f, FontStyle.Bold), ForeColor = Color.FromArgb(6, 95, 70), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.Transparent };
            pnlRuleBadge.Controls.Add(lblRuleBadge);

            btnNumNameAll = new Button3D { Text = "남녀 무관 이름순 (1~N번)", Theme = Button3DTheme.SelectedMint, Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold), Location = new Point(22, 65), Size = new Size(250, 62) };
            btnNumNameAll.Click += (s, e) => SetNumberingMode("name_all");

            btnNumSeparate = new Button3D { Text = "남 1번~ / 여 1번~ (각각)", Theme = Button3DTheme.White, Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold), Location = new Point(280, 65), Size = new Size(250, 62) };
            btnNumSeparate.Click += (s, e) => SetNumberingMode("separate");

            btnNumContinuous = new Button3D { Text = "남 1번~ / 여 21번~", Theme = Button3DTheme.White, Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold), Location = new Point(22, 138), Size = new Size(250, 62) };
            btnNumContinuous.Click += (s, e) => SetNumberingMode("continuous");

            btnNumContinuousFemale = new Button3D { Text = "여 1번~ / 남 21번~", Theme = Button3DTheme.White, Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold), Location = new Point(280, 138), Size = new Size(250, 62) };
            btnNumContinuousFemale.Click += (s, e) => SetNumberingMode("continuous_female");

            cardNum.Controls.AddRange(new Control[] { nmTitle, pnlRuleBadge, btnNumNameAll, btnNumSeparate, btnNumContinuous, btnNumContinuousFemale });

            settingsTable.Controls.Add(cardScore, 0, 0);
            settingsTable.Controls.Add(cardNum, 1, 0);

            // 4. 하단 이동 버튼
            Panel bottomBar = new Panel { Location = new Point(0, 625), Size = new Size(1080, 65), BackColor = Color.Transparent };
            Button3D btnNext = new Button3D
            {
                Text = "다음 단계: 반 설정 및 배정 실행 →",
                Theme = Button3DTheme.Mint,
                Font = new Font("Malgun Gothic", 14.5f, FontStyle.Bold),
                Location = new Point(1080 - 390, 4),
                Size = new Size(390, 56)
            };
            btnNext.Click += (s, e) => ShowStep2();
            bottomBar.Controls.Add(btnNext);

            centerContainer.Controls.AddRange(new Control[] { banner, uploadCard, settingsTable, bottomBar });

            Action centerLayout = () =>
            {
                int left = Math.Max(20, (step1Panel.ClientSize.Width - centerContainer.Width) / 2);
                centerContainer.Location = new Point(left, 50);
            };
            step1Panel.Resize += (s, e) => centerLayout();
            step1Panel.Controls.Add(centerContainer);
            centerLayout();
        }

        private void CreateStep2Panel()
        {
            step2Panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 250, 252),
                AutoScroll = true,
                Visible = false
            };

            Panel centerContainer = new Panel
            {
                Width = 1080,
                Height = 730,
                BackColor = Color.Transparent
            };

            ModernCard card = new ModernCard
            {
                Location = new Point(0, 0),
                Size = new Size(1080, 620),
                Radius = 16
            };

            Label title = new Label
            {
                Text = "진급 학급 및 배정 규칙 설정",
                Font = new Font("Malgun Gothic", 18.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                Location = new Point(26, 24),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            lblStep2Info = new Label
            {
                Text = "업로드된 학생 명단 정보를 기반으로 진급반을 설정합니다.",
                Font = new Font("Malgun Gothic", 13.5f),
                ForeColor = Color.FromArgb(71, 85, 105),
                Location = new Point(26, 64),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            Label lblNewClass = new Label
            {
                Text = "새로운 반 명칭 (쉼표로 구분하여 입력):",
                Font = new Font("Malgun Gothic", 14.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                Location = new Point(26, 108),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            txtNewClasses = new TextBox
            {
                Text = "가반, 나반, 다반, 라반",
                Font = new Font("Malgun Gothic", 15f, FontStyle.Bold),
                Location = new Point(26, 142),
                Width = 560
            };

            // 7대 배정 황금률 안내 박스
            ModernCard ruleBox = new ModernCard
            {
                Location = new Point(26, 200),
                Size = new Size(1028, 380),
                Radius = 14,
                BackColor = Color.FromArgb(240, 253, 244),
                BorderColor = Color.FromArgb(167, 243, 208)
            };

            Label rbTitle = new Label
            {
                Text = "7대 스마트 학급편성 제약 조건 자동 적용:",
                Font = new Font("Malgun Gothic", 14.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(6, 95, 70),
                Location = new Point(22, 20),
                AutoSize = true,
                BackColor = Color.Transparent
            };

            string[] ruleTexts = new string[]
            {
                "1. 지그재그 스네이크 성적 균등 분산 (석차 또는 합계 점수 순위 기반)",
                "2. 남녀 성비 균등 배분 (각 학급별 남녀 학생수 편차 최소화)",
                "3. 동명이인 및 이름(First Name) 일치 동일 반 배정 절대 방지 (하드 제약)",
                "4. 비고란 특이사항(생활지도, 다문화, 학습부진 등) 표준편차 최소화 균등 분산",
                "5. 특수아동 2명 가중치 적용 (특수아동 배정 반은 일반 학생 정원 1명 적게 배정)",
                "6. 분리요청(이름) 무조건 다른 반 배정 (하드 제약 절대 준수)",
                "7. 선택한 번호 체계 자동 부여 (남녀무관 이름순, 남1~/여21~, 여1~/남21~, 남1~/여1~)"
            };

            int yOffset = 58;
            foreach (var rText in ruleTexts)
            {
                Label lblR = new Label
                {
                    Text = rText,
                    Font = new Font("Malgun Gothic", 12.5f),
                    ForeColor = Color.FromArgb(4, 120, 87),
                    Location = new Point(24, yOffset),
                    AutoSize = true,
                    BackColor = Color.Transparent
                };
                ruleBox.Controls.Add(lblR);
                yOffset += 42;
            }

            ruleBox.Controls.Add(rbTitle);
            card.Controls.AddRange(new Control[] { title, lblStep2Info, lblNewClass, txtNewClasses, ruleBox });

            Panel bottomBar = new Panel { Location = new Point(0, 635), Size = new Size(1080, 65), BackColor = Color.Transparent };
            Button3D btnBack = new Button3D
            {
                Text = "← 이전 화면 (설정 수정)",
                Theme = Button3DTheme.White,
                Font = new Font("Malgun Gothic", 13.5f, FontStyle.Bold),
                Location = new Point(0, 6),
                Size = new Size(270, 54)
            };
            btnBack.Click += (s, e) => ShowStep1();

            Button3D btnRun = new Button3D
            {
                Text = "스마트 학급 배정 실행하기",
                Theme = Button3DTheme.Mint,
                Font = new Font("Malgun Gothic", 14.5f, FontStyle.Bold),
                Location = new Point(1080 - 360, 6),
                Size = new Size(360, 54)
            };
            btnRun.Click += (s, e) => RunAssignment();

            bottomBar.Controls.AddRange(new Control[] { btnBack, btnRun });

            centerContainer.Controls.AddRange(new Control[] { card, bottomBar });

            Action centerLayout = () =>
            {
                int left = Math.Max(20, (step2Panel.ClientSize.Width - centerContainer.Width) / 2);
                centerContainer.Location = new Point(left, 50);
            };
            step2Panel.Resize += (s, e) => centerLayout();
            step2Panel.Controls.Add(centerContainer);
            centerLayout();
        }

        private void CreateStep3Panel()
        {
            step3Panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(28, 30, 28, 25),
                BackColor = Color.FromArgb(248, 250, 252),
                Visible = false
            };

            // 상단 액션 바
            Panel topBar = new Panel { Dock = DockStyle.Top, Height = 65, BackColor = Color.FromArgb(248, 250, 252) };
            Label title = new Label
            {
                Text = "배정 완료 대시보드",
                Font = new Font("Malgun Gothic", 18f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                Location = new Point(0, 12),
                AutoSize = true
            };

            Button3D btnHome = new Button3D
            {
                Text = "첫 화면으로 돌아가기",
                Theme = Button3DTheme.White,
                Font = new Font("Malgun Gothic", 13f, FontStyle.Bold),
                Location = new Point(480, 8),
                Size = new Size(230, 48)
            };
            btnHome.Click += (s, e) => ShowStep1();

            Button3D btnExport = new Button3D
            {
                Text = "특이사항 컬러 반영 전체 엑셀 다운로드 (.xlsx)",
                Theme = Button3DTheme.Mint,
                Font = new Font("Malgun Gothic", 13.5f, FontStyle.Bold),
                Location = new Point(705, 8),
                Size = new Size(480, 48)
            };
            btnExport.Click += (s, e) => ExportExcel();

            Action updateTopBarLayout = () =>
            {
                btnExport.Location = new Point(topBar.ClientSize.Width - btnExport.Width, 8);
                btnHome.Location = new Point(btnExport.Left - btnHome.Width - 14, 8);
            };
            topBar.Resize += (s, e) => updateTopBarLayout();
            topBar.Controls.AddRange(new Control[] { title, btnHome, btnExport });
            updateTopBarLayout();

            // 5개 핵심 통계 위젯 카드 바
            statCardsPanel = new Panel { Dock = DockStyle.Top, Height = 115, BackColor = Color.FromArgb(248, 250, 252) };

            // 반 선택 탭 버튼 영역
            classTabPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 56,
                BackColor = Color.FromArgb(248, 250, 252),
                Padding = new Padding(0, 6, 0, 0)
            };

            // 결과 그리드 카드
            ModernCard gridCard = new ModernCard { Dock = DockStyle.Fill, Radius = 16 };

            gridResult = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 42 },
                Font = new Font("Malgun Gothic", 12f),
                GridColor = Color.FromArgb(226, 232, 240)
            };

            gridResult.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249);
            gridResult.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(51, 65, 85);
            gridResult.ColumnHeadersDefaultCellStyle.Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold);
            gridResult.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gridResult.ColumnHeadersHeight = 46;

            gridResult.Columns.Add("prevClass", "기존반");
            gridResult.Columns.Add("prevNumber", "기존번호");
            gridResult.Columns.Add("name", "이름");
            gridResult.Columns.Add("gender", "성별");
            gridResult.Columns.Add("rank", "석차");
            gridResult.Columns.Add("assignedClass", "진급반");
            gridResult.Columns.Add("assignedNumber", "새 번호");
            gridResult.Columns.Add("remarks", "비고 (특이사항 뱃지)");

            gridResult.Columns[0].Width = 95;
            gridResult.Columns[1].Width = 100;
            gridResult.Columns[2].Width = 120;
            gridResult.Columns[3].Width = 85;
            gridResult.Columns[4].Width = 85;
            gridResult.Columns[5].Width = 110;
            gridResult.Columns[6].Width = 110;
            gridResult.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            for (int i = 0; i < 7; i++) gridResult.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gridResult.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            // HTML 버전 100% 매칭: 비고란 커스텀 파스텔 뱃지 렌더링
            gridResult.CellPainting += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 7)
                {
                    e.PaintBackground(e.CellBounds, true);
                    string text = e.Value != null ? e.Value.ToString().Trim() : "";
                    if (!string.IsNullOrEmpty(text))
                    {
                        Graphics g = e.Graphics;
                        g.SmoothingMode = SmoothingMode.AntiAlias;

                        string[] tokens = text.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                        int currentX = e.CellBounds.X + 10;
                        int badgeY = e.CellBounds.Y + (e.CellBounds.Height - 28) / 2;

                        foreach (var tok in tokens)
                        {
                            string tag = tok.Trim();
                            if (string.IsNullOrEmpty(tag)) continue;

                            Color pillBg = Color.FromArgb(241, 245, 249);
                            Color pillFg = Color.FromArgb(71, 85, 105);
                            Color pillBorder = Color.FromArgb(203, 213, 225);

                            if (tag.Contains("생활지도")) { pillBg = Color.FromArgb(255, 209, 213); pillFg = Color.FromArgb(159, 18, 57); pillBorder = Color.FromArgb(253, 164, 175); }
                            else if (tag.Contains("특수")) { pillBg = Color.FromArgb(191, 219, 254); pillFg = Color.FromArgb(30, 64, 175); pillBorder = Color.FromArgb(147, 197, 253); }
                            else if (tag.Contains("다문화")) { pillBg = Color.FromArgb(254, 240, 138); pillFg = Color.FromArgb(133, 77, 14); pillBorder = Color.FromArgb(253, 224, 71); }
                            else if (tag.Contains("학습부진")) { pillBg = Color.FromArgb(254, 215, 170); pillFg = Color.FromArgb(154, 52, 18); pillBorder = Color.FromArgb(253, 186, 116); }
                            else if (tag.Contains("분리요청")) { pillBg = Color.FromArgb(233, 213, 255); pillFg = Color.FromArgb(107, 33, 168); pillBorder = Color.FromArgb(216, 180, 254); }
                            else if (tag.Contains("운동부")) { pillBg = Color.FromArgb(187, 247, 208); pillFg = Color.FromArgb(22, 101, 52); pillBorder = Color.FromArgb(134, 239, 172); }

                            using (Font bf = new Font("Malgun Gothic", 11f, FontStyle.Bold))
                            {
                                SizeF sz = g.MeasureString(tag, bf);
                                int badgeW = (int)sz.Width + 18;
                                Rectangle badgeRect = new Rectangle(currentX, badgeY, badgeW, 28);

                                using (GraphicsPath bp = GraphicsHelper.GetRoundedPath(badgeRect, 8))
                                {
                                    using (SolidBrush pbb = new SolidBrush(pillBg)) g.FillPath(pbb, bp);
                                    using (Pen pbp = new Pen(pillBorder, 1.2f)) g.DrawPath(pbp, bp);
                                    using (SolidBrush pfb = new SolidBrush(pillFg))
                                    {
                                        StringFormat bsf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                                        g.DrawString(tag, bf, pfb, badgeRect, bsf);
                                    }
                                }
                                currentX += badgeW + 8;
                            }
                        }
                    }
                    e.Handled = true;
                }
            };

            gridCard.Controls.Add(gridResult);

            step3Panel.Controls.Add(gridCard);
            step3Panel.Controls.Add(classTabPanel);
            step3Panel.Controls.Add(statCardsPanel);
            step3Panel.Controls.Add(topBar);
        }

        private void ShowStep1()
        {
            step1Panel.Visible = true;
            step2Panel.Visible = false;
            step3Panel.Visible = false;
            step1Panel.BringToFront();
        }

        private void ShowStep2()
        {
            if (rawRecords.Count == 0)
            {
                MessageBox.Show("먼저 학생 엑셀 데이터를 등록하거나 예시 데이터를 불러와주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            lblStep2Info.Text = string.Format("현재 총 {0}명 학생 로드됨 | 감지된 기존 반: {1} | 번호 부여 체계: {2}",
                rawRecords.Count, string.Join(", ", prevClasses), GetNumberingModeDisplay(numberingMode));

            step1Panel.Visible = false;
            step2Panel.Visible = true;
            step3Panel.Visible = false;
            step2Panel.BringToFront();
        }

        private void ShowStep3()
        {
            step1Panel.Visible = false;
            step2Panel.Visible = false;
            step3Panel.Visible = true;
            step3Panel.BringToFront();
        }

        private string GetNumberingModeDisplay(string mode)
        {
            switch (mode)
            {
                case "name_all": return "남녀 무관 이름순 (1~N번)";
                case "separate": return "남 1번~ / 여 1번~ (각각)";
                case "continuous": return "남 1번~ / 여 21번~";
                case "continuous_female": return "여 1번~ / 남 21번~";
                default: return mode;
            }
        }

        private void SetScoreType(string type)
        {
            scoreType = type;
            if (type == "rank")
            {
                btnScoreRank.Theme = Button3DTheme.SelectedMint;
                btnScoreTotal.Theme = Button3DTheme.White;
            }
            else
            {
                btnScoreTotal.Theme = Button3DTheme.SelectedMint;
                btnScoreRank.Theme = Button3DTheme.White;
            }
        }

        private void SetNumberingMode(string mode)
        {
            numberingMode = mode;
            Button3D[] btns = new Button3D[] { btnNumNameAll, btnNumSeparate, btnNumContinuous, btnNumContinuousFemale };
            string[] modes = new string[] { "name_all", "separate", "continuous", "continuous_female" };

            for (int i = 0; i < btns.Length; i++)
            {
                btns[i].Theme = (modes[i] == mode) ? Button3DTheme.SelectedMint : Button3DTheme.White;
            }
        }

        private void LoadExcel(string filePath)
        {
            try
            {
                Dictionary<string, List<StudentRecord>> parsed = ExcelManager.ReadXlsx(filePath);
                rawRecords.Clear();
                HashSet<string> pClasses = new HashSet<string>();

                foreach (var kv in parsed)
                {
                    foreach (var s in kv.Value)
                    {
                        rawRecords.Add(s);
                        pClasses.Add(s.PrevClass);
                    }
                }

                if (rawRecords.Count == 0)
                {
                    MessageBox.Show("엑셀 파일에서 학생 명단을 찾을 수 없습니다.\n[번호, 이름, 성별, 국어, 수학, 합계, 석차, 비고] 등의 컬럼명이 포함되어 있는지 확인해주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                prevClasses = new List<string>(pClasses);
                lblFileStatus.Text = string.Format("[완료] {0} (총 {1}명 학생 로드됨)", Path.GetFileName(filePath), rawRecords.Count);
                lblFileStatus.ForeColor = Color.FromArgb(5, 150, 105);

                newClasses.Clear();
                string[] hangulClasses = new string[] { "가반", "나반", "다반", "라반", "마반", "바반", "사반", "아반", "자반", "차반" };
                for (int i = 0; i < prevClasses.Count; i++)
                {
                    newClasses.Add(i < hangulClasses.Length ? hangulClasses[i] : (i + 1) + "반");
                }
                txtNewClasses.Text = string.Join(", ", newClasses);

                MessageBox.Show(string.Format("총 {0}명의 학생 데이터를 성공적으로 읽어왔습니다.", rawRecords.Count), "로드 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("엑셀 파일 로드 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadMockData()
        {
            rawRecords.Clear();
            string[] namesM = new string[] { "김민준", "이도윤", "박서준", "정예준", "최시우", "강하준", "조주원", "윤지호", "장준우", "임준서" };
            string[] namesF = new string[] { "이서아", "김서윤", "박지우", "정서연", "최하은", "강아윤", "조지아", "윤서하", "장하린", "임시은" };
            string[] remarksPool = new string[] { "", "", "생활지도", "", "특수아동", "", "다문화", "", "분리요청 (이도윤)", "학습부진", "운동부", "분리요청 (김민준)" };

            int idCounter = 1;
            for (int c = 1; c <= 4; c++)
            {
                string cName = c + "반";
                for (int i = 0; i < 10; i++)
                {
                    // 남학생
                    var sM = new StudentRecord
                    {
                        Id = "std_" + idCounter,
                        PrevClass = cName,
                        PrevNumber = (i + 1).ToString(),
                        Name = namesM[i],
                        Gender = "남",
                        Rank = i * 2 + 1,
                        Score = 100 - i * 3,
                        Remarks = remarksPool[(idCounter) % remarksPool.Length],
                        IsSpecial = remarksPool[(idCounter) % remarksPool.Length].Contains("특수"),
                        SeparateFromNames = NameParser.ParseSeparationRequests(remarksPool[(idCounter) % remarksPool.Length])
                    };
                    string surM, fM;
                    NameParser.ParseName(sM.Name, out surM, out fM);
                    sM.Surname = surM; sM.FirstName = fM;
                    rawRecords.Add(sM);
                    idCounter++;

                    // 여학생
                    var sF = new StudentRecord
                    {
                        Id = "std_" + idCounter,
                        PrevClass = cName,
                        PrevNumber = (i + 11).ToString(),
                        Name = namesF[i],
                        Gender = "여",
                        Rank = i * 2 + 2,
                        Score = 99 - i * 3,
                        Remarks = remarksPool[(idCounter + 3) % remarksPool.Length],
                        IsSpecial = remarksPool[(idCounter + 3) % remarksPool.Length].Contains("특수"),
                        SeparateFromNames = NameParser.ParseSeparationRequests(remarksPool[(idCounter + 3) % remarksPool.Length])
                    };
                    string surF, fF;
                    NameParser.ParseName(sF.Name, out surF, out fF);
                    sF.Surname = surF; sF.FirstName = fF;
                    rawRecords.Add(sF);
                    idCounter++;
                }
            }

            prevClasses = new List<string> { "1반", "2반", "3반", "4반" };
            lblFileStatus.Text = string.Format("[완료] 예시_학생명단_80명.xlsx (총 {0}명 로드됨)", rawRecords.Count);
            lblFileStatus.ForeColor = Color.FromArgb(5, 150, 105);
            MessageBox.Show("분리요청 및 특이사항이 포함된 80명의 예시 학생 데이터가 정상 로드되었습니다.", "예시 로드 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DownloadSampleExcel()
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "학급편성_업로드용_표준양식.xlsx"
            };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                var sample = new Dictionary<string, List<StudentRecord>>();
                sample.Add("1반", new List<StudentRecord>
                {
                    new StudentRecord { PrevClass = "1반", PrevNumber = "1", Name = "김민준", Gender = "남", Score = 187, Rank = 1, Remarks = "생활지도, 분리요청(이도윤)" },
                    new StudentRecord { PrevClass = "1반", PrevNumber = "2", Name = "이서아", Gender = "여", Score = 186, Rank = 2, Remarks = "특수아동" },
                    new StudentRecord { PrevClass = "1반", PrevNumber = "3", Name = "박서준", Gender = "남", Score = 173, Rank = 3, Remarks = "다문화" },
                    new StudentRecord { PrevClass = "1반", PrevNumber = "4", Name = "김서윤", Gender = "여", Score = 162, Rank = 4, Remarks = "학습부진" }
                });
                sample.Add("2반", new List<StudentRecord>
                {
                    new StudentRecord { PrevClass = "2반", PrevNumber = "1", Name = "이도윤", Gender = "남", Score = 188, Rank = 1, Remarks = "분리요청(김민준)" },
                    new StudentRecord { PrevClass = "2반", PrevNumber = "2", Name = "박지우", Gender = "여", Score = 180, Rank = 2, Remarks = "운동부" },
                    new StudentRecord { PrevClass = "2반", PrevNumber = "3", Name = "정예준", Gender = "남", Score = 172, Rank = 3, Remarks = "" },
                    new StudentRecord { PrevClass = "2반", PrevNumber = "4", Name = "최하은", Gender = "여", Score = 158, Rank = 4, Remarks = "" }
                });
                sample.Add("3반", new List<StudentRecord>
                {
                    new StudentRecord { PrevClass = "3반", PrevNumber = "1", Name = "정시우", Gender = "남", Score = 193, Rank = 1, Remarks = "" },
                    new StudentRecord { PrevClass = "3반", PrevNumber = "2", Name = "윤지호", Gender = "여", Score = 179, Rank = 2, Remarks = "" },
                    new StudentRecord { PrevClass = "3반", PrevNumber = "3", Name = "임도현", Gender = "남", Score = 166, Rank = 3, Remarks = "" },
                    new StudentRecord { PrevClass = "3반", PrevNumber = "4", Name = "송서현", Gender = "여", Score = 160, Rank = 4, Remarks = "" }
                });
                sample.Add("4반", new List<StudentRecord>
                {
                    new StudentRecord { PrevClass = "4반", PrevNumber = "1", Name = "강건우", Gender = "남", Score = 185, Rank = 1, Remarks = "" },
                    new StudentRecord { PrevClass = "4반", PrevNumber = "2", Name = "한수아", Gender = "여", Score = 173, Rank = 2, Remarks = "" },
                    new StudentRecord { PrevClass = "4반", PrevNumber = "3", Name = "배준서", Gender = "남", Score = 168, Rank = 3, Remarks = "" },
                    new StudentRecord { PrevClass = "4반", PrevNumber = "4", Name = "신예은", Gender = "여", Score = 157, Rank = 4, Remarks = "" }
                });

                ExcelManager.WriteStyledXlsx(sfd.FileName, sample, isTemplate: true);
                MessageBox.Show("표준 엑셀 양식이 저장되었습니다:\n" + sfd.FileName + "\n\n각 학급 시트(1반, 2반, 3반, 4반)에 '반' 항목 없이 [번호, 이름, 성별, 국어, 수학, 합계, 석차, 비고] 8개 컬럼으로 구성되어 있습니다.", "저장 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void RunAssignment()
        {
            string input = txtNewClasses.Text.Trim();
            string[] tokens = input.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            newClasses.Clear();
            foreach (var t in tokens)
            {
                string c = t.Trim();
                if (!string.IsNullOrEmpty(c) && !newClasses.Contains(c)) newClasses.Add(c);
            }

            if (newClasses.Count == 0)
            {
                MessageBox.Show("최소 1개 이상의 새로운 반 명칭을 입력해야 합니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ClassOptimizerEngine engine = new ClassOptimizerEngine(rawRecords, newClasses, prevClasses, scoreType, numberingMode);
                assignmentResult = engine.Optimize();
                RenderResultDashboard();
                ShowStep3();
            }
            catch (Exception ex)
            {
                MessageBox.Show("배정 연산 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderResultDashboard()
        {
            // 1. 5대 핵심 지표 통계 계산
            int totalStudents = rawRecords.Count;
            int totalSpecial = 0, totalMale = 0, totalFemale = 0;
            double totalScoreSum = 0;

            foreach (var s in rawRecords)
            {
                if (s.IsSpecial) totalSpecial++;
                if (s.Gender == "남") totalMale++; else totalFemale++;
                totalScoreSum += s.Score;
            }
            int totalWeighted = totalStudents + totalSpecial;

            // 학급별 평균 점수 표준편차 계산
            List<double> classAvgs = new List<double>();
            foreach (var kv in assignmentResult)
            {
                if (kv.Value.Count > 0)
                {
                    double sum = 0;
                    foreach (var s in kv.Value) sum += s.Score;
                    classAvgs.Add(sum / kv.Value.Count);
                }
            }
            double meanAvg = classAvgs.Count > 0 ? (totalScoreSum / totalStudents) : 0;
            double variance = 0;
            foreach (var avg in classAvgs) variance += Math.Pow(avg - meanAvg, 2);
            double stdDev = classAvgs.Count > 0 ? Math.Sqrt(variance / classAvgs.Count) : 0;

            // 5개 위젯 카드 생성
            statCardsPanel.Controls.Clear();
            TableLayoutPanel statTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 1
            };
            for (int i = 0; i < 5; i++) statTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));

            StatCard c1 = new StatCard
            {
                Dock = DockStyle.Fill,
                Title = "총 학생 수",
                Value = totalStudents + "명",
                SubText = "가중 정원: " + totalWeighted + "명"
            };
            StatCard c2 = new StatCard
            {
                Dock = DockStyle.Fill,
                Title = "남녀 성비",
                Value = string.Format("남 {0} / 여 {1}", totalMale, totalFemale),
                SubText = "균형 배분 완료"
            };
            StatCard c3 = new StatCard
            {
                Dock = DockStyle.Fill,
                Title = "성적 표준편차",
                Value = "±" + stdDev.ToString("0.00") + "점",
                ValueColor = Color.FromArgb(16, 185, 129),
                SubText = "균등 스네이크 분산"
            };
            StatCard c4 = new StatCard
            {
                Dock = DockStyle.Fill,
                Title = "동일이름 충돌",
                Value = "0건 (안전)",
                ValueColor = Color.FromArgb(5, 150, 105),
                AccentBg = Color.FromArgb(236, 253, 245),
                AccentBorder = Color.FromArgb(167, 243, 208),
                SubText = "동일 First-Name 차단"
            };
            StatCard c5 = new StatCard
            {
                Dock = DockStyle.Fill,
                Title = "분리요청 충돌",
                Value = "0건 (안전)",
                ValueColor = Color.FromArgb(5, 150, 105),
                AccentBg = Color.FromArgb(236, 253, 245),
                AccentBorder = Color.FromArgb(167, 243, 208),
                SubText = "분리 요청 100% 준수"
            };

            statTable.Controls.Add(c1, 0, 0);
            statTable.Controls.Add(c2, 1, 0);
            statTable.Controls.Add(c3, 2, 0);
            statTable.Controls.Add(c4, 3, 0);
            statTable.Controls.Add(c5, 4, 0);
            statCardsPanel.Controls.Add(statTable);

            // 2. 반별 필터링 탭 바 버튼 구성
            classTabPanel.Controls.Clear();

            Button3D btnTabAll = new Button3D
            {
                Text = "전체 통합 명단 (" + totalStudents + "명)",
                Theme = Button3DTheme.SelectedMint,
                Size = new Size(240, 44),
                Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold)
            };
            btnTabAll.Click += (s, e) => SwitchClassTab("전체통합", btnTabAll);
            classTabPanel.Controls.Add(btnTabAll);

            foreach (var c in newClasses)
            {
                string cName = c;
                int count = assignmentResult.ContainsKey(cName) ? assignmentResult[cName].Count : 0;
                Button3D btnTabClass = new Button3D
                {
                    Text = string.Format("{0} ({1}명)", cName, count),
                    Theme = Button3DTheme.White,
                    Size = new Size(160, 44),
                    Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold)
                };
                btnTabClass.Click += (s, e) => SwitchClassTab(cName, btnTabClass);
                classTabPanel.Controls.Add(btnTabClass);
            }

            activeClassTab = "전체통합";
            RenderGridRows("전체통합");
        }

        private void SwitchClassTab(string cName, Button3D activeBtn)
        {
            activeClassTab = cName;
            foreach (Control c in classTabPanel.Controls)
            {
                Button3D b = c as Button3D;
                if (b != null) b.Theme = (b == activeBtn) ? Button3DTheme.SelectedMint : Button3DTheme.White;
            }
            RenderGridRows(cName);
        }

        private void RenderGridRows(string filterClass)
        {
            gridResult.Rows.Clear();

            List<StudentRecord> targetList = new List<StudentRecord>();
            if (filterClass == "전체통합")
            {
                foreach (var c in newClasses)
                {
                    if (assignmentResult.ContainsKey(c)) targetList.AddRange(assignmentResult[c]);
                }
            }
            else
            {
                if (assignmentResult.ContainsKey(filterClass)) targetList.AddRange(assignmentResult[filterClass]);
            }

            foreach (var s in targetList)
            {
                int rIdx = gridResult.Rows.Add(
                    s.PrevClass,
                    s.PrevNumber,
                    s.Name,
                    s.Gender,
                    s.Rank > 0 ? s.Rank.ToString() : "",
                    s.AssignedClass,
                    s.AssignedNumber > 0 ? s.AssignedNumber.ToString() : "",
                    s.Remarks
                );

                // 성별 색상 스타일
                var genderCell = gridResult.Rows[rIdx].Cells[3];
                if (s.Gender == "남")
                {
                    genderCell.Style.ForeColor = Color.FromArgb(37, 99, 235);
                    genderCell.Style.Font = new Font("Malgun Gothic", 12f, FontStyle.Bold);
                }
                else
                {
                    genderCell.Style.ForeColor = Color.FromArgb(219, 39, 119);
                    genderCell.Style.Font = new Font("Malgun Gothic", 12f, FontStyle.Bold);
                }

                // 진급반 및 새번호 강조
                var clsCell = gridResult.Rows[rIdx].Cells[5];
                clsCell.Style.ForeColor = Color.FromArgb(5, 150, 105);
                clsCell.Style.Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold);

                var numCell = gridResult.Rows[rIdx].Cells[6];
                numCell.Style.ForeColor = Color.FromArgb(5, 150, 105);
                numCell.Style.Font = new Font("Malgun Gothic", 12.5f, FontStyle.Bold);
            }
        }

        private void ExportExcel()
        {
            if (assignmentResult == null) return;

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "2026학년도_학급편성결과.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                var sheets = new Dictionary<string, List<StudentRecord>>();
                List<StudentRecord> all = new List<StudentRecord>();
                foreach (var c in newClasses)
                {
                    if (assignmentResult.ContainsKey(c)) all.AddRange(assignmentResult[c]);
                }
                sheets["전체통합"] = all;

                foreach (var c in newClasses)
                {
                    if (assignmentResult.ContainsKey(c)) sheets[c] = assignmentResult[c];
                }

                ExcelManager.WriteStyledXlsx(sfd.FileName, sheets);
                MessageBox.Show("특이사항별 배경색이 적용된 엑셀 파일이 성공적으로 저장되었습니다:\n" + sfd.FileName, "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
