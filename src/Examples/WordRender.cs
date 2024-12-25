using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Examples
{
    public static class WordRender
    {
        static void ReplaceParserTag(this OpenXmlElement elem, 
            Dictionary<string, string> data, 
            string matchPattern = @"\[\$(?<n>\w+)\$\]")
        {
            var pool = new List<Run>();
            var matchText = string.Empty;
            var hiliteRuns = elem.Descendants<Run>() 
                .Where(o => o.RunProperties?.Elements<Highlight>().Any() ?? false).ToList();

            foreach (var run in hiliteRuns)
            {
                var t = run.InnerText;
                if (t.StartsWith("["))
                {
                    pool = new List<Run> { run };
                    matchText = t;
                }
                else
                {
                    matchText += t;
                    pool.Add(run);
                }
                if (t.EndsWith("]"))
                {
                    var m = Regex.Match(matchText, matchPattern);
                    if (m.Success && data.ContainsKey(m.Groups["n"].Value))
                    {
                        var firstRun = pool.First();
                        firstRun.RemoveAllChildren<Text>();
                        firstRun.RunProperties.RemoveAllChildren<Highlight>();
                        var newText = data[m.Groups["n"].Value];
                        var firstLine = true;
                        foreach (var line in Regex.Split(newText, @"\\n"))
                        {
                            if (firstLine) firstLine = false;
                            else firstRun.Append(new Break());
                            firstRun.Append(new Text(line));
                        }
                        pool.Skip(1).ToList().ForEach(o => o.Remove());
                    }
                }
            }
        }

        static void RenderTableRow(this OpenXmlElement elem, Dictionary<string, IEnumerable<object>> tableData)
        {
            const string matchPattern = @"\[\#(?<k>\w+)[.](?<n>\w+)\#\]";
            elem.Descendants<TableRow>().Where(o => o.InnerText.Contains("[#"))
                .ToList()
                .ForEach(tmplRow =>
                {
                    var m = Regex.Match(tmplRow.InnerText, matchPattern);
                    if (m.Success && tableData.ContainsKey(m.Groups["k"].Value))
                    {
                        var list =
                            JsonConvert.DeserializeObject<Dictionary<string,string>[]>(
                                JsonConvert.SerializeObject(tableData[m.Groups["k"].Value])
                            );
                        foreach (var item in list)
                        {
                            var cloneRow = (TableRow)tmplRow.Clone();
                            cloneRow.ReplaceParserTag(item, matchPattern);
                            tmplRow.Parent.Append(cloneRow);
                        }
                        tmplRow.Remove();
                    }
                });
        }

        public static byte[] GenerateDocx(byte[] template, Dictionary<string, string> data, Dictionary<string, IEnumerable<object>> tableData = null)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(template, 0, template.Length);
                using (var docx = WordprocessingDocument.Open(ms, true))
                {
                    docx.MainDocumentPart.HeaderParts.ToList().ForEach(hdr =>
                    {
                        hdr.Header.ReplaceParserTag(data);
                    });
                    docx.MainDocumentPart.FooterParts.ToList().ForEach(ftr =>
                    {
                        ftr.Footer.ReplaceParserTag(data);
                    });
                    docx.MainDocumentPart.Document.Body.ReplaceParserTag(data);
                    if (tableData != null)
                        docx.MainDocumentPart.Document.Body.RenderTableRow(tableData);
                    docx.Save();
                }
                return ms.ToArray();
            }
        }

        public static byte[] CreateNew()
        {
            using (var ms = new MemoryStream())
            {
                using (var docx = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = docx.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    mainPart.Document.AppendChild(new Body());
                    docx.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(new Text("Hello"))));
                    docx.Save();
                }
                return ms.ToArray();
            }
        }
    }
}
