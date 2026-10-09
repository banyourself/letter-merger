using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace LetterMerger
{
    public sealed class TextChange
    {
        public string Before, After;
    }

    public sealed class TextRepairResult
    {
        public string Text;
        public bool Unresolved;
        public List<TextChange> Changes = new List<TextChange>();
    }

    public static class TextRepair
    {
        static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);
        static readonly Dictionary<char, byte> ExtraBytes = LegacyBytes();

        static Dictionary<char, byte> LegacyBytes()
        {
            var result = new Dictionary<char, byte>();
            string characters = "\u20ac\u0081\u201a\u0192\u201e\u2026\u2020\u2021\u02c6\u2030\u0160\u2039\u0152\u008d\u017d\u008f\u0090\u2018\u2019\u201c\u201d\u2022\u2013\u2014\u02dc\u2122\u0161\u203a\u0153\u009d\u017e\u0178";
            for (int i = 0; i < characters.Length; i++)
                if (characters[i] > 255) result[characters[i]] = (byte)(i + 128);
            return result;
        }

        static bool Byte(char value, out byte result)
        {
            if (value <= 255)
            {
                result = (byte)value;
                return true;
            }
            return ExtraBytes.TryGetValue(value, out result);
        }

        public static TextRepairResult Fix(string source)
        {
            var result = new TextRepairResult { Text = source ?? "" };
            for (int pass = 0; pass < 3; pass++)
            {
                string current = result.Text;
                var output = new StringBuilder(current.Length);
                int before = result.Changes.Count;
                for (int i = 0; i < current.Length; i++)
                {
                    byte first;
                    int length = Byte(current[i], out first) ? first >= 194 && first <= 207 ? 2 : first == 226 ? 3 : first == 240 ? 4 : 0 : 0;
                    if (length == 0 || i + length > current.Length)
                    {
                        output.Append(current[i]);
                        continue;
                    }
                    var bytes = new byte[length];
                    bytes[0] = first;
                    bool valid = true;
                    for (int j = 1; j < length; j++)
                        if (!Byte(current[i + j], out bytes[j]) || bytes[j] < 128 || bytes[j] > 191)
                        {
                            valid = false;
                            break;
                        }
                    string decoded = null;
                    if (valid)
                        try { decoded = Utf8.GetString(bytes); }
                        catch (DecoderFallbackException) { }
                    if (decoded == null)
                    {
                        output.Append(current[i]);
                        continue;
                    }
                    result.Changes.Add(new TextChange { Before = current.Substring(i, length), After = decoded });
                    output.Append(decoded);
                    i += length - 1;
                }
                result.Text = output.ToString();
                if (result.Changes.Count == before) break;
            }
            result.Unresolved = Regex.IsMatch(result.Text, "[\uFFFD\u0080-\u009F]|\u00e2[\u20ac\u0080]|\u00f0\u0178");
            return result;
        }

        public static void Prepare(Student student, bool enabled)
        {
            student.SourceMessage = student.Message;
            student.EncodingRepairEnabled = enabled;
            var result = Fix(student.SourceMessage);
            student.TextChanges = result.Changes;
            if (enabled) student.Message = result.Text;
            if (!enabled && result.Changes.Count > 0)
                student.TextIssues.Add("TEXT_ENCODING_REVIEW: recognizable encoding errors remain because repair is off.");
            if (result.Unresolved)
                student.TextIssues.Add("TEXT_ENCODING_REVIEW: incomplete or lost characters require checking the original response; no replacement was guessed.");
            foreach (var field in new[] { Tuple.Create("student name", student.Name), Tuple.Create("scholarship", student.Scholarship) })
            {
                var identity = Fix(field.Item2);
                if (identity.Changes.Count > 0 || identity.Unresolved)
                    student.TextIssues.Add("TEXT_ENCODING_REVIEW: " + field.Item1 + " may contain an encoding error; its exact wording is preserved.");
            }
        }

        public static IEnumerable<string[]> ReportRows(IEnumerable<Student> students)
        {
            foreach (var student in students)
            {
                foreach (var group in student.TextChanges.GroupBy(c => c.Before + "\u0000" + c.After))
                {
                    var change = group.First();
                    yield return new[] { student.Record.ToString(), student.DataRow.ToString(), student.Name, student.Scholarship, "Thank-you text", student.EncodingRepairEnabled ? "REPAIRED" : "REPAIR_DISABLED", change.Before, change.After, group.Count().ToString(), "Complete response retained; source CSV unchanged." };
                }
                foreach (string issue in student.TextIssues)
                    yield return new[] { student.Record.ToString(), student.DataRow.ToString(), student.Name, student.Scholarship, "Review", "NEEDS_REVIEW", "", "", "", issue };
            }
        }

        public static readonly string[] Headers = { "CSVRecord", "DataRow", "StudentName", "Scholarship", "Field", "Status", "Before", "After", "Occurrences", "Details" };
    }

    public sealed class MergeOutput
    {
        public string Path;
        public List<Student> Cards;
        public double Fold;
    }

    public sealed class MergeBatchResult
    {
        public List<MergeOutput> Outputs = new List<MergeOutput>();
        public List<string> Errors = new List<string>();
    }

    public static class BatchMerge
    {
        public const string MainFile = "Merged_Letters.docx", MissingFile = "Merged_Letters_Missing.docx";

        public static MergeBatchResult Build(string templatePath, List<Student> students, string batch, string mode, Overlay overlay, bool separateMissing, Action<string> progress)
        {
            var result = new MergeBatchResult();
            foreach (var student in students)
            {
                student.OutputFile = "";
                student.PhotoAdded = false;
                student.Card = 0;
            }
            if (students.Count == 0) return result;
            string expectedHash = Core.Hash(templatePath);
            Core.RejectLink(batch);
            string working = System.IO.Path.Combine(batch, "Work_" + Guid.NewGuid().ToString("N").Substring(0, 12));
            Core.RejectLink(working);
            Directory.CreateDirectory(working);
            try
            {
                var missing = separateMissing ? students.Where(s => String.IsNullOrEmpty(s.Photo)).ToList() : new List<Student>();
                var main = separateMissing ? students.Where(s => !String.IsNullOrEmpty(s.Photo)).ToList() : students.ToList();
                BuildGroup(templatePath, expectedHash, main, missing, batch, working, MainFile, mode, overlay, separateMissing, result, progress);
                if (separateMissing)
                {
                    missing = missing.OrderBy(s => s.DataRow).ToList();
                    BuildGroup(templatePath, expectedHash, missing, null, batch, working, MissingFile, mode, overlay, false, result, progress);
                }
            }
            finally
            {
                try
                {
                    Core.RejectLink(working);
                    Directory.Delete(working, true);
                }
                catch (Exception e) { result.Errors.Add("Temporary merge files could not be removed: " + e.GetBaseException().Message); }
            }
            return result;
        }

        static void BuildGroup(string templatePath, string expectedHash, List<Student> cards, List<Student> missing, string batch, string working, string file, string mode, Overlay overlay, bool regroup, MergeBatchResult result, Action<string> progress)
        {
            try
            {
                while (cards.Count > 0)
                {
                    progress("Generating " + cards.Count + " cards in " + file + ".");
                    var template = Core.ReadTemplate(templatePath);
                    if (template.Hash != expectedHash) throw new InvalidDataException("Template changed between merged files.");
                    string temporary = System.IO.Path.Combine(working, Guid.NewGuid().ToString("N").Substring(0, 12) + ".docx");
                    string validation = temporary + ".validation.txt";
                    Core.Build(template, cards, temporary, mode, overlay, validation);
                    var failed = regroup ? cards.Where(s => !s.PhotoAdded).ToList() : new List<Student>();
                    if (failed.Count > 0)
                    {
                        foreach (var student in failed)
                        {
                            student.Photo = "";
                            student.Card = 0;
                            missing.Add(student);
                            cards.Remove(student);
                        }
                        File.Delete(temporary);
                        if (File.Exists(validation)) File.Delete(validation);
                        continue;
                    }
                    string output = System.IO.Path.GetFullPath(System.IO.Path.Combine(batch, file));
                    File.Move(temporary, output);
                    foreach (var student in cards) student.OutputFile = output;
                    result.Outputs.Add(new MergeOutput { Path = output, Cards = cards.ToList(), Fold = template.Fold });
                    if (File.Exists(validation))
                        try
                        {
                            string note = file == MissingFile || regroup ? "DocumentValidation_" + System.IO.Path.GetFileNameWithoutExtension(file) + ".txt" : "DocumentValidation.txt";
                            File.Move(validation, System.IO.Path.Combine(batch, note));
                        }
                        catch (Exception e) { result.Errors.Add(file + " validation note: " + e.GetBaseException().Message); }
                    return;
                }
            }
            catch (Exception e) { result.Errors.Add(file + ": " + e.GetBaseException().Message); }
        }
    }
}
