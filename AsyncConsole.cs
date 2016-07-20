using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using Newtonsoft.Json;
using Microsoft.Scripting.Utils;
using System.Xml.Linq;

namespace WordAnalysis
{
    using System.Threading;

    public partial class AsyncConsole : Form
    {
        public AsyncConsole()
        {
            InitializeComponent();
        }

        public IEnumerable<string> EnumerableFS(string folder, Func<string, bool> where = null)
        {
            folder = TrimPath(folder);
            if (new DirectoryInfo(folder).Exists)
            {
                string[] folders = null;
                try
                {
                    folders = Directory.GetDirectories(folder);
                }
                catch (Exception)
                {
                }
                Func<string, string[]> getFiles = Directory.GetFiles;
                Func<string, string[]> tryGetFiles = getFiles.Try;
                if (folders != null)
                {
                    foreach (string subfolder in folders)
                    {
                        var f = TrimPath(subfolder);
                        foreach (string ssfolder in this.EnumerableFS(f, @where))
                        {
                            if (@where != null && @where(ssfolder)) yield return ssfolder;
                        }
                        if (Directory.Exists(f))
                        {
                            string[] filePaths = tryGetFiles(f);
                            if (filePaths != null)
                            {
                                foreach (string s in filePaths)
                                {
                                    if (@where != null && @where(s)) yield return s;
                                }
                            }
                        }
                    }
                }
                string[] files = tryGetFiles(folder);
                if (files != null)
                {
                    foreach (string file in files)
                    {
                        if (where != null && where(file)) yield return file;
                    }
                }
            }
        }

        private static string TrimPath(string folder)
        {
            if (folder.Length > 248)
            {
                folder = folder.Substring(0, 248);
            }
            return folder;
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(
                o =>
                    {
                        Invoke(
                            new MethodInvoker(
                                () =>
                                    {
                                        Clipboard.SetText(
                                            SetText(
                                                EnumerableFS(
                                                    @"D:\MyConfiguration\lzy13870\Desktop\sent",
                                                    p => p.EndsWith(".doc") || p.EndsWith(".docx"))
                                                    .Take(2)
                                                    .Select(p => Parse(OpenWordXml(p)))
                                                    .ToList()
                                                    .JoinStrings(Environment.NewLine)));
                                    }));
                        Application.Exit();
                    });
        }

        public string SetText(string txt)
        {
            Invoke(new MethodInvoker(() => this.txtOutput.Text = txt));
            return txt;
        }

        public static string WordDocument(string path)
        {
            if (!File.Exists(path))
            {
                return string.Empty;
            }
            //Process.Start("taskkill", " /f /t /im WINWORD.EXE");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = null;
            app.Visible = true;
            document = Open(path, app);
            StringBuilder buf = new StringBuilder();
            object oMissing = global::System.Reflection.Missing.Value;

            var xmlText =
                document.Content.XML.ToString();
            document.Close();
            app.Quit();

            //File.WriteAllText("temp.xml", xmlText);
            //Process.Start("temp.xml");
            //buf.AppendLine(xml.SelectNodes("//tbl").ToEnumerable().Take(1).FirstOrDefault().OuterXml);
            buf.AppendLine(path);
            buf.AppendLine(Parse(xmlText));
            Clipboard.SetText(buf.ToString());
            return buf.ToString();
        }

        private static string Parse(string xmlText)
        {
            if (xmlText == null)
            {
                return string.Empty;
            }
            var xml = Xml(
                xmlText.Replace("w:", string.Empty).Replace("wx:", string.Empty).Replace("wsp:", string.Empty));
            StringBuilder buf = new StringBuilder();
            buf.AppendLine(
                xml.Select("/wordDocument/body/sect/sub-section/sub-section/p[1]/r/t")
                    .Select(p => p.InnerText)
                    .JoinStrings());
            buf.AppendLine();
            buf.AppendLine(
                xml.Select("//tr")
                    .Select(
                        tr =>
                        tr.Select("tc")
                            .Select(
                                tc =>
                                tc.Select("p")
                                    .Select(
                                        p =>
                                        p.Select("r")
                                            .Select(r => r.Select("t").Select(t => t.InnerText).JoinStrings("\t"))
                                            .JoinStrings())
                                    .JoinStrings("|"))
                            .JoinStrings("|"))
                    .JoinStrings(Environment.NewLine));
            var s = buf.ToString();



            string[] lines = s.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            return
                Str(
                    lines.OfType<string>()
                        .Select(ParseLine).Where(p=>p!=null).Select(Json).JoinStrings(Environment.NewLine),
                    Environment.NewLine,
                    s
                    //Environment.NewLine,
                    //xml.Select("//tr").Select(tr => PrettyXml(tr.OuterXml)).JoinStrings(Environment.NewLine)
                    );
        }

        public static object ParseLine(string line)
        {
            string resourceTypePattern = "(门票|用餐|住宿|导游|保险|其他|交通){1}";
            string unitPattern = @"((元|辆|人|餐|间|车|天|晚|桌|场)+\/*)+";
            string numberPattern = @"([0-9]+[\,\.\+\-\*\/\^\=\s]*)+";
            string resourceTypeColumnPattern = @"\|*[0-9]{1,3}、*(门票|用餐|住宿|导游|保险|其他|交通){1}\|*";
            string pricePattern = Str(numberPattern, unitPattern);
            string fastText =
                line.Matches(resourceTypeColumnPattern).JoinStrings("\t").Matches(resourceTypePattern).JoinStrings("\t");
            line = line.ReplaceRegex("服务标准");
            if (fastText.IsNotEmpty())
            {
                line = line.Trim('|');
                var row = line.Split(new[] { '|' }).ToList();
                var resourceType = row.Get(0).MatchesJoinTrim(resourceTypePattern);
                var resourceName = row.Get(1).Trim();
                if (row.Count > 5)
                {
                    var subPrice =
                        row.Skip(2)
                            .Take(row.Count - 4)
                            .Where(item => item.Matches(pricePattern).Take(1).JoinStrings().Trim().IsNotEmpty());
                    var subPriceStart = row.IndexOf(subPrice.FirstOrDefault());
                    var subPriceCount = subPrice.Count();
                    var subPriceEnd = subPriceStart + subPriceCount;
                    return subPrice.Select(
                        item =>
                            {
                                var idx = (row.IndexOf(item) - subPriceStart);
                                var count = ParseMathForm(row.Get(subPriceEnd + idx).MatchesJoinTrim(numberPattern));
                                var total =
                                    row.Get(subPriceEnd + subPriceCount + idx)
                                        .Matches(numberPattern)
                                        .IsNull(new[] { "0" })
                                        .Last();
                                var price =
                                    item.Matches(pricePattern).Take(1).JoinStrings().MatchesJoinTrim(numberPattern);
                                var unit = item.Matches(unitPattern).HashSet().JoinStrings();
                                var days = item.Matches(@"\*\d+").Get(0).MatchesJoinTrim(@"\d+");

                                if (total.IsEmpty())
                                {
                                    total = count;
                                    count = string.Empty;
                                }
                                return
                                    new
                                        {
                                            Id = -1,
                                            QuoteNoteId = 0,
                                            ResourcesType = GetResourcesType(resourceType),
                                            ResourcesTypeText = resourceType,
                                            ResourcesName = resourceName,
                                            ResourcesPrice = price,
                                            UnitName = unit,
                                            Days = days.MatchesJoinTrim(numberPattern),
                                            ResourcesCount = count,
                                            ResourcesTotalPrice = total
                                        };
                            });
                }
                var resourcePrice =
                    row.Get(2)
                        .Matches(pricePattern)
                        .Take(1)
                        .JoinStrings()
                        .MatchesJoinTrim(numberPattern)
                        .IsNull(resourceName.MatchesJoinTrim(pricePattern).MatchesJoinTrim(numberPattern));
                var resourceUnit =
                    row.Get(2)
                        .Matches(unitPattern)
                        .HashSet()
                        .JoinStrings()
                        .IsNull(resourceName.Matches(unitPattern).HashSet().JoinStrings());
                var resourceDays =
                    row.Get(2)
                        .Matches(@"\*\d+")
                        .Get(0)
                        .MatchesJoinTrim(@"\d+")
                        .IsNull(resourceName.Matches(@"\*\d+").Get(0).MatchesJoinTrim(@"\d+"));
                var resourceCount =
                    ParseMathForm(
                        row.Get(3)
                            .Matches(numberPattern)
                            .FirstOrDefault()
                            .IsNull(
                                resourceName.MatchesJoinTrim(numberPattern)
                                    .Split("*".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).JoinStrings(" "))
                            .IsNull("0"));
                var resourceTotal =
                    row.Get(4)
                        .Matches(numberPattern)
                        .IsNull(resourceName.Matches(numberPattern))
                        .LastOrDefault().IsNull("0");
                if (resourceTotal.IsEmpty())
                {
                    resourceTotal = resourceCount;
                    resourceCount = string.Empty;
                }
                return new[]
                           {
                               new
                                   {
                                       Id = -1,
                                       QuoteNoteId = 0,
                                       ResourcesType = GetResourcesType(resourceType),
                                       ResourcesTypeText = resourceType,
                                       ResourcesName = resourceName,
                                       ResourcesPrice = resourcePrice,
                                       UnitName = resourceUnit,
                                       Days = resourceDays.MatchesJoinTrim(@"\d+"),
                                       ResourcesCount = resourceCount,
                                       ResourcesTotalPrice = resourceTotal
                                   }
                           };
            }
            else if (line.MatchesJoinTrim(@"\|*(D|第)*[0-9]+天*\：*\|*").IsNotEmpty())
            {
                string dayPattern = @"\|*\s*(D|第)+\s*[0-9]+\s*天*\s*(\：|\:)*\s*\|*";
                string timePattern =
                    @"\|+(([0-9]+\s*(\:|\：)+\s*[0-9]+)+(\s*(\-|\—|\—\—)+\s*[0-9]+\s*(\:|\：)+\s*[0-9]+)*)+";
                var quote = new
                                {
                                    QuoteNoteList = line.MatchesSplit(dayPattern).Select(
                                        day =>
                                            {
                                                var dayDesc =
                                                    day.MatchesJoinTrim(@"([\u4e00-\u9fa5]+(\s*(\-|\—|\—\—)+\s*[\u4e00-\u9fa5]+)+\s*\|+)+")
                                                        .Trim('|');
                                                var dayIndexDesc = day.MatchesJoin(dayPattern, ",");
                                                return
                                                    new
                                                        {
                                                            JourneyDay =
                                                                new { JourneyTitle = dayDesc, Days = dayIndexDesc.MatchesJoinTrim(@"\d+") },
                                                            QuoteJourneyDayList = day.MatchesSplit(timePattern).Select(
                                                                time =>
                                                                    {
                                                                        var timeDesc = time.MatchesJoinTrim(timePattern);
                                                                        var resourceType = ParseResourceType(time);
                                                                        time = time.ReplaceRegex(timeDesc);
                                                                        return
                                                                            new
                                                                                {
                                                                                    Time = timeDesc.Trim('|'),
                                                                                    JourneyType = GetJourneyType(ParseResourceType(time)),
                                                                                    JourneyTypeText = ParseResourceType(time),
                                                                                    Contents = time.ReplaceRegex(timeDesc).Trim('|')
                                                                                };
                                                                    }).ToList(),
                                                            GalleryList = new[] { new { } }
                                                        };
                                            }).ToList()
                                };
                return quote;
            }
            return null;
        }

        private static int GetJourneyType(string journeyType)
        {
            if (journeyType == "酒店") return 7;
            if (journeyType == "自由活动") return 1;
            if (journeyType == "购物") return 6;
            if (journeyType == "景点") return 4;
            if (journeyType == "用餐"|| journeyType == "餐饮") return 3;
            if (journeyType == "大巴"|| journeyType == "交通") return 5;
            return 0;
        }

        private static int GetResourcesType(string resourceType)
        {
            if (resourceType == "大巴" || resourceType == "交通") return 1;
            if (resourceType == "酒店" || resourceType == "住宿") return 2;
            if (resourceType == "景点" || resourceType == "门票") return 3;
            if (resourceType == "导游") return 4;
            if (resourceType == "餐饮" || resourceType == "用餐") return 5;
            if (resourceType == "保险") return 7;
            if (resourceType == "其他") return 6;
            return 0;
        }

        public static string ParseMathForm(string input)
        {
            if (input==null)
            {
                return null;
            }
            var funcs = new Dictionary<string, Func<string, string>>()
                            {
                                {
                                    " ",
                                    str =>
                                    str.SplitEx(" ").Take(1)
                                        .Select(ParseMathForm)
                                        .Select(double.Parse)
                                        .ToString()
                                },
                                {
                                    "=",
                                    str =>
                                    str.SplitEx("=").Take(1)
                                        .Select(ParseMathForm)
                                        .Select(double.Parse)
                                        .ToString()
                                },
                                {
                                    "+",
                                    str =>
                                    str.SplitEx("+")
                                        .Select(ParseMathForm)
                                        .Select(double.Parse)
                                        .Aggregate((a, b) => a + b)
                                        .ToString()
                                },
                                {
                                    "-",
                                    str =>
                                    str.SplitEx("-")
                                        .Select(ParseMathForm)
                                        .Select(double.Parse)
                                        .Aggregate((a, b) => a - b)
                                        .ToString()
                                },
                                {
                                    "/",
                                    str =>
                                    str.SplitEx("/")
                                        .Select(ParseMathForm)
                                        .Select(double.Parse)
                                        .Aggregate((a, b) => a / b)
                                        .ToString()
                                },
                                {
                                    "*",
                                    str =>
                                    str.SplitEx("*")
                                        .Select(ParseMathForm)
                                        .Select(double.Parse)
                                        .Aggregate((a, b) => a * b)
                                        .ToString()
                                }
                            };
            var key = funcs.Keys.FirstOrDefault(k => input.Contains(k));
            if (key.IsNotEmpty())
            {
                return funcs[key](input);
            }
            return input;
        }

        public static string PrettyXml(string xml)
        {
            var stringBuilder = new StringBuilder();

            var element = XElement.Parse(xml);

            var settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;
            settings.Indent = true;
            settings.NewLineOnAttributes = true;

            using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
            {
                element.Save(xmlWriter);
            }

            return stringBuilder.ToString();
        }

        private static string ParseResourceType(string time)
        {
            var map = new Dictionary<string, string>()
                          {
                              { "(出发|返回|乘|指定地点|大巴)+", "交通" },
                              { "(游|风景区)+", "景点" },
                              { "(酒店)+", "酒店" },
                              { "(餐)+", "餐饮" },
                          };
            return
                map.Keys.Select(key => time.MatchesJoinTrim(key).IsNotEmpty() ? map[key] : string.Empty)
                    .Where(p => p.IsNotEmpty())
                    .Take(1)
                    .JoinStrings();
        }

        public static XmlElement Xml(string xml)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(xml);
                return document.DocumentElement;
            }
            catch (Exception exception)
            {
                Exceptions.Add(exception);
            }
            return null;
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Range> Loop(Microsoft.Office.Interop.Word.Sentences sentences)
        {
            for (int i = 1; i < sentences.Count; i++)
            {
                yield return sentences[i];
            }
        }

        public static Microsoft.Office.Interop.Word.Document Open(string path, Microsoft.Office.Interop.Word.Application app)
        {
            object file = path;
            object unknow = Type.Missing;
            Microsoft.Office.Interop.Word.Document document = app.Documents.Open(
                ref file,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow,
                ref unknow);
            return document;
        }

        public static IEnumerable<int> Range(int start, int end)
        {
            for (int i = start; i <= end; i++)
            {
                yield return i;
            }
        }

        public static string Repeat(string text, int n)
        {
            return Str(Enumerable.Select(Range(0, n), i => text).ToList().ToArray());
        }

        public static string PP(Microsoft.Office.Interop.Word.Range range, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder(Str(Repeat("\t", n), "Range", range.Text));

            buf.AppendLine(Str(Repeat("\t", current), "Tables",
                Str(
                    Try(
                        range,
                        r =>
                        Loop(r.Tables)
                            .Select(table => Str(PP(table, current)))
                            .ToList()
                            .ToArray()))));
            buf.AppendLine(Str(Repeat("\t", current), "Bookmarks",
                Str(
                    Loop(range.Bookmarks)
                        .Select(bookmark => Str(PP(bookmark, current)))
                        .ToList()
                        .ToArray())));

            buf.AppendLine(Str(Repeat("\t", current), "Words",
                Str(
                    Try(
                        range,
                        r =>
                        Loop(r.Words)
                            .Select(bookmark => Str(PP(bookmark, current)))
                            .ToList()
                            .ToArray()))));
            range.Select();
            buf.AppendLine(Str(Repeat("\t", current), "Cells",
                Str(
                    Try(
                        range,
                        r =>
                        Loop(r.Cells)
                            .Select(bookmark => Str(PP(bookmark, current)))
                            .ToList()
                            .ToArray()))));
            return buf.ToString();
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Range> Loop(Microsoft.Office.Interop.Word.Words words)
        {
            for (int i = 0; i < words.Count; i++)
            {
                yield return words[i];
            }
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Cell> Loop(Microsoft.Office.Interop.Word.Cells cells)
        {
            for (int i = 1; i < cells.Count; i++)
            {
                yield return cells[i];
            }
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Bookmark> Loop(Microsoft.Office.Interop.Word.Bookmarks bookmarks)
        {
            for (int j = 1; j < bookmarks.Count; j++)
            {
                yield return bookmarks[j];
            }
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Table> Loop(Microsoft.Office.Interop.Word.Tables tables)
        {
            for (int j = 1; j < tables.Count; j++)
            {
                yield return tables[j];
            }
        }

        public static List<Exception> Exceptions = new List<Exception>();

        public static TOut Try<TIn, TOut>(TIn source, Func<TIn, TOut> fn) where TOut : class where TIn : class
        {
            try
            {
                return fn(source);
            }
            catch (Exception exception)
            {
                Exceptions.Add(exception);
            }

            return null;
        }

        public static string DrawLine(int current, Microsoft.Office.Interop.Word.Cell cell)
        {
            return Str(Repeat("\t", current), PP(cell, current));
        }

        public static string PP(Microsoft.Office.Interop.Word.Cell cell, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.Append(Str("Cell", Repeat("\t", current), PP(cell.Range, current)));
            return buf.ToString();
        }

        public static string PP(Microsoft.Office.Interop.Word.Bookmark bookmark, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            buf.AppendLine(Str("Bookmark", Repeat("\t", current), bookmark.Name, PP(bookmark.Range, current)));
            return buf.ToString();
        }

        public static string PP(Microsoft.Office.Interop.Word.Table table, int n)
        {
            int current = n + 1;
            return Str(Loop(table.Rows).Select(r => Str("Row", Repeat("\t", current), PP(r, current))).ToList().ToArray());
        }

        public static IEnumerable<Microsoft.Office.Interop.Word.Row> Loop(Microsoft.Office.Interop.Word.Rows rows)
        {
            for (int k = 1; k < rows.Count; k++)
            {
                yield return rows[k];
            }
        }

        public static string PP(Microsoft.Office.Interop.Word.Row row, int n)
        {
            int current = n + 1;
            StringBuilder buf = new StringBuilder();
            if (row.Cells.Count > 0)
            {
                for (int l = 1; l < row.Cells.Count; l++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Word.Cell cell = row.Cells[l];
                        buf.AppendLine(Str(Repeat("\t", current), PP(cell.Range, current)));
                    }
                    catch (Exception exception)
                    {
                        buf.AppendLine(exception.Message);
                        throw exception;
                    }
                }
            }

            return buf.ToString();
        }

        public static string Str(params object[] strings)
        {
            if (strings != null)
            {
                return string.Join(string.Empty, strings);
            }

            return string.Empty;
        }

        public static string Json(object obj)
        {
            try
            {
                return JsonConvert.SerializeObject(obj);
            }
            catch (Exception exception)
            {
                MessageBox.Show(obj.ToString(), exception.Message);
            }
            return string.Empty;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            var result = this.openFileDialogDefault.ShowDialog();
            if (result == DialogResult.OK)
            {
                var fileName = this.openFileDialogDefault.FileName;
                var xmlText = OpenWordXml(fileName);
                this.txtOutput.Text = Str(fileName, Environment.NewLine, Parse(xmlText));
            }
        }

        public void openFileDialog(Action<string> fn)
        {
            var result = this.openFileDialogDefault.ShowDialog();
            if (result == DialogResult.OK)
            {
                fn(openFileDialogDefault.FileName);
            }
        }
        public void openDirDialog(Action<string> fn)
        {
            var result = this.folderBrowserDialogDefault.ShowDialog();
            if (result == DialogResult.OK)
            {
                fn(folderBrowserDialogDefault.SelectedPath);
            }
        }

        private static string OpenWordXml(string fileName)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            var document = Open(fileName, app);
            Func<Microsoft.Office.Interop.Word.Document, string> getXml = doc => doc.Content.XML;
            var xmlText = getXml.Try(document);
            document.Close();
            app.Quit();
            return xmlText;
        }

        private void textBoxFolder_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBoxFolder_DoubleClick(object sender, EventArgs e)
        {
            openDirDialog(
                path => ThreadPool.QueueUserWorkItem(
                    o =>
                    {
                        SetText(
                            EnumerableFS(path, p => p.EndsWith(".doc") || p.EndsWith(".docx"))
                                //.Take(2)
                                .Select(p => Parse(OpenWordXml(p)))
                                .ToList()
                                .JoinStrings(Environment.NewLine));
                    }));
        }
    }
}
