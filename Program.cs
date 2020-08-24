using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TextCopy;

namespace Projects
{
    class Program
    {
        static void Main(string[] args)
        {
            File.Copy("程式異動申請清單.tpl.docx", "程式異動申請清單.docx", true);

/*
.vscodeignore	.vscodeignore	Added	1	0
CHANGELOG.md	.md	Modified	43	3
README.md	.md	Modified	89	5
package.json	.json	Modified	4	2
snippets/go.json	.json	Modified	26	6
*/
            // 讀取剪貼簿內容
            var text = ClipboardService.GetText();

            #if DEBUG
            text = @".vscodeignore	.vscodeignore	Added	1	0
CHANGELOG.md	.md	Modified	43	3
a/b/c/README.md	.md	Modified	89	5
package.json	.json	Modified	4	2
snippets/go.json	.json	Modified	26	6";
            #endif

            var records = new List<GitDiff>();
            foreach (var item in Regex.Split(text, "\r\n|\r|\n"))
            {
                var line = item.Trim();

                if (String.IsNullOrEmpty(line))
                {
                    continue;
                }

                var fields = line.Split(new[] { '\t' });
                records.Add(new GitDiff()
                {
                    File = fields[0],
                    Extension = fields[1],
                    Action = fields[2],
                    LinesAdded = fields[3],
                    LinesRemoved = fields[4]
                });
            }

            using (WordprocessingDocument doc = WordprocessingDocument.Open("程式異動申請清單.docx", true))
            {
                // Find the first table in the document.
                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                var rows = table.Elements<TableRow>().Count();

                TableRow secondRow = table.Elements<TableRow>().ElementAt(1).CloneNode(true) as TableRow;
                TableRow lastRow = table.Elements<TableRow>().ElementAt(rows - 1).CloneNode(true) as TableRow;

                for (int i = rows - 1; i > 0; i--)
                {
                    table.Elements<TableRow>().ElementAt(i).Remove();
                }

                for (int i = 0; i < records.Count; i++)
                {
                    var row = secondRow.CloneNode(true) as TableRow;
                    SetCell(row, 0, String.Format("{0:D2}", i));
                    SetCell(row, 1, Path.GetFileName(records[i].File));
                    SetCell(row, 2, Path.GetDirectoryName(records[i].File));
                    SetCell(row, 3, GetActionCode(records[i].Action));
                    table.Append(row);
                }

                table.Append(lastRow.CloneNode(true) as TableRow);
            }
        }

        private static string GetActionCode(string action)
        {
            switch (action)
            {
                case "Added": return "A";
                case "Modified": return "U";
                case "Deleted": return "D";
                default: return "C";
            }
        }

        private static void SetCell(TableRow row, int cellIndex, string text)
        {
            TableCell cell = row.Elements<TableCell>().ElementAt(cellIndex);

            var all_text = cell.Descendants<Text>();
            foreach (var item in all_text)
            {
                item.Remove();
            }

            cell.LastChild.Append(new Run(new Text(text)));
        }
    }
}
