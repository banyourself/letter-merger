using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace LetterMerger
{
    public sealed class LayoutRow
    {
        public string Card = "", Record = "", DataRow = "", Name = "", Scholarship = "", Pages = "", StartPage = "", EndPage = "", FirstY = "", LastY = "", Font = "", Words = "", Status = "", Details = "";
        public string OutputFile = "";
        public string[] Values()
        {
            return new[]
            {
                Card,
                Record,
                DataRow,
                Name,
                Scholarship,
                Pages,
                StartPage,
                EndPage,
                FirstY,
                LastY,
                Font,
                Words,
                Status,
                Details,
                OutputFile
            };
        }
    }

    public static class WordLayout
    {
        public static readonly string[] Headers =
        {
            "Card",
            "CSVRecord",
            "DataRow",
            "StudentName",
            "Scholarship",
            "Pages",
            "StartPage",
            "EndPage",
            "FirstMessageY",
            "LastMessageY",
            "BodyFont",
            "ResponseWords",
            "Status",
            "Details",
            "OutputFile"
        };
        static void Release(object x)
        {
            if (x != null && Marshal.IsComObject(x))
                try
                {
                    Marshal.ReleaseComObject(x);
                }
                catch
                {
                }
        }

        static double Point(dynamic doc, int start, int end, int kind)
        {
            dynamic r = null, w = null;
            try
            {
                r = doc.Range(start, end);
                w = doc.ActiveWindow;
                w.ScrollIntoView(r, true);
                return Convert.ToDouble(r.Information[kind]);
            }
            finally
            {
                Release(w);
                Release(r);
            }
        }

        static LayoutRow Row(Student s)
        {
            return new LayoutRow
            {
                Card = s.Card.ToString(),
                Record = s.Record.ToString(),
                DataRow = s.DataRow.ToString(),
                Name = s.Name,
                Scholarship = s.Scholarship,
                Font = s.Font.ToString(CultureInfo.InvariantCulture),
                Words = s.Words.ToString()
            };
        }

        public static List<LayoutRow> Check(string path, List<Student> cards, double fold, bool pdf, string batch, Action<string> progress)
        {
            var report = new List<LayoutRow>();
            dynamic word = null, doc = null, documents = null, sections = null, bookmarks = null;
            string stage = "WORD_START";
            bool good = true;
            try
            {
                Core.RejectLink(path);
                Core.Validate(path, cards);
                string checkedHash = Core.Hash(path);
                if (Environment.OSVersion.Platform != PlatformID.Win32NT)
                    throw new PlatformNotSupportedException("Desktop Microsoft Word checks require Windows.");
                Type type = Type.GetTypeFromProgID("Word.Application");
                if (type == null)
                    throw new InvalidOperationException("Desktop Microsoft Word is unavailable.");
                word = Activator.CreateInstance(type);
                word.Visible = true;
                word.DisplayAlerts = 0;
                word.AutomationSecurity = 3;
                documents = word.Documents;
                stage = "WORD_OPEN";
                if (Core.Hash(path) != checkedHash)
                    throw new InvalidDataException("The merged document changed before the Word layout check.");
                doc = documents.Open(path, false, true, false);
                stage = "WORD_LAYOUT";
                dynamic window = null, view = null;
                try
                {
                    window = doc.ActiveWindow;
                    view = window.View;
                    view.Type = 3;
                }
                finally
                {
                    Release(view);
                    Release(window);
                }

                doc.Repaginate();
                sections = doc.Sections;
                bookmarks = doc.Bookmarks;
                int total = Convert.ToInt32(doc.ComputeStatistics(2)), sectionCount = Convert.ToInt32(sections.Count);
                good = total == cards.Count && sectionCount == cards.Count;
                foreach (var s in cards)
                {
                    progress("Checking Word layout for card " + s.Card + " of " + cards.Count + ".");
                    var row = Row(s);
                    var issues = Core.ContentIssues(s);
                    dynamic section = null, range = null, bookmark = null, message = null, setup = null;
                    try
                    {
                        if (s.Card > sectionCount)
                            throw new InvalidDataException("Corresponding Word section is missing.");
                        section = sections.Item(s.Card);
                        range = section.Range;
                        if (range.End > range.Start)
                            range.End = range.End - 1;
                        int pages = Convert.ToInt32(range.ComputeStatistics(2));
                        double start = Point(doc, (int)range.Start, (int)range.Start + 1, 3), end = Point(doc, Math.Max((int)range.Start, (int)range.End - 1), (int)range.End, 3);
                        row.Pages = pages.ToString();
                        row.StartPage = start.ToString();
                        row.EndPage = end.ToString();
                        if (pages != 1 || start != end)
                            issues.Add("OVERFLOW: card occupies " + pages + " pages, from " + start + " to " + end + ".");
                        string mark = "LM_Message_" + s.Card.ToString("D4");
                        if (!bookmarks.Exists(mark))
                            throw new InvalidDataException("Response bookmark is missing.");
                        bookmark = bookmarks.Item(mark);
                        message = bookmark.Range;
                        if (!Core.Normalize((string)message.Text).Contains(Core.Normalize(s.Message)))
                            issues.Add("MESSAGE_MISMATCH: full supplied response could not be verified.");
                        double first = Point(doc, (int)message.Start, (int)message.Start + 1, 6), last = Point(doc, Math.Max((int)message.Start, (int)message.End - 1), (int)message.End, 6);
                        row.FirstY = first.ToString(CultureInfo.InvariantCulture);
                        row.LastY = last.ToString(CultureInfo.InvariantCulture);
                        if (first < 0 || last < 0)
                            issues.Add("POSITION_UNCHECKED: inspect the fold line and bottom manually.");
                        else
                        {
                            if (first < fold + 8)
                                issues.Add("FOLD_OVERLAP: message begins too close to/above the fold line.");
                            setup = section.PageSetup;
                            double bottom = (double)setup.PageHeight - (double)setup.BottomMargin;
                            if (last + s.Font > bottom + 1)
                                issues.Add("BOTTOM_OVERFLOW: last line extends past the printable body.");
                        }

                        foreach (string prefix in new[]
                        {
                            "LM_Name_",
                            "LM_Scholar_"
                        }

                        )
                        {
                            dynamic bm = null, r = null;
                            try
                            {
                                bm = bookmarks.Item(prefix + s.Card.ToString("D4"));
                                r = bm.Range;
                                double y = Point(doc, Math.Max((int)r.Start, (int)r.End - 1), (int)r.End, 6);
                                if (y < 0)
                                    issues.Add("POSITION_UNCHECKED: inspect upper text at fold line.");
                                else if (y + (prefix == "LM_Name_" ? 16 : 14) > fold - 4)
                                    issues.Add("UPPER_TEXT_NEAR_FOLD: review name/scholarship without changing their wording.");
                            }
                            finally
                            {
                                Release(r);
                                Release(bm);
                            }
                        }

                        row.Status = issues.Count == 0 ? "LAYOUT_OK" : "NEEDS_REVIEW";
                        row.Details = issues.Count == 0 ? "One page; full response verified; text positions checked." : String.Join(" ", issues);
                        if (issues.Count > 0)
                            good = false;
                    }
                    catch (Exception e)
                    {
                        row.Status = "LAYOUT_ERROR";
                        row.Details = e.GetBaseException().Message + " " + String.Join(" ", issues);
                        good = false;
                    }
                    finally
                    {
                        Release(setup);
                        Release(message);
                        Release(bookmark);
                        Release(range);
                        Release(section);
                    }

                    report.Add(row);
                }

                if (total != cards.Count || sectionCount != cards.Count)
                    report.Add(new LayoutRow { Pages = total.ToString(), Status = "BATCH_LAYOUT_MISMATCH", Details = total + " pages and " + sectionCount + " sections for " + cards.Count + " cards." });
                if (pdf && good)
                    try
                    {
                        string pdfPath = Path.Combine(batch, Path.GetFileNameWithoutExtension(path) + ".pdf");
                        Core.RejectLink(pdfPath);
                        if (File.Exists(pdfPath))
                            throw new IOException("A PDF already exists at the output path.");
                        doc.ExportAsFixedFormat(pdfPath, 17);
                    }
                    catch (Exception e)
                    {
                        report.Add(new LayoutRow { Status = "PDF_ERROR", Details = e.GetBaseException().Message });
                    }
            }
            catch (Exception e)
            {
                report.Clear();
                foreach (var s in cards)
                {
                    var row = Row(s);
                    row.Status = stage == "WORD_OPEN" ? "WORD_OPEN_ERROR" : "LAYOUT_UNCHECKED";
                    row.Details = e.GetBaseException().Message + " " + String.Join(" ", Core.ContentIssues(s));
                    report.Add(row);
                }
            }
            finally
            {
                Release(bookmarks);
                Release(sections);
                if (doc != null)
                {
                    try
                    {
                        doc.Close(false);
                    }
                    catch (Exception e)
                    {
                        report.Add(new LayoutRow { Status = "WORD_CLEANUP_WARNING", Details = e.Message });
                    }

                    Release(doc);
                }

                Release(documents);
                if (word != null)
                {
                    try
                    {
                        word.Quit();
                    }
                    catch (Exception e)
                    {
                        report.Add(new LayoutRow { Status = "WORD_CLEANUP_WARNING", Details = e.Message });
                    }

                    Release(word);
                }
            }

            foreach (var row in report) row.OutputFile = Path.GetFullPath(path);
            return report;
        }
    }
}
