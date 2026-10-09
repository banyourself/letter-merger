using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Xml;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LetterMerger
{
    public static class Tests
    {
        static List<string> results;
        static void Check(string name, Action action)
        {
            action();
            results.Add("PASS: " + name);
            Console.WriteLine("PASS: " + name);
        }

        static void Assert(bool value, string message)
        {
            if (!value)
                throw new InvalidOperationException(message);
        }

        static void Reject(Action action)
        {
            bool rejected = false;
            try
            {
                action();
            }
            catch
            {
                rejected = true;
            }

            Assert(rejected, "Unsafe/invalid input was accepted.");
        }

        static byte[] Entry(string path, string name)
        {
            using (var z = ZipFile.OpenRead(path))
            using (var s = z.GetEntry(name).Open())
            using (var m = new MemoryStream())
            {
                s.CopyTo(m);
                return m.ToArray();
            }
        }

        static string Patch(string source, string output, string part, byte[] bytes)
        {
            Core.WritePackage(source, output, new Dictionary<string, byte[]> { { part, bytes } });
            return output;
        }

        static void ImageFile(string path, int width, int height, ImageFormat format)
        {
            using (var b = new Bitmap(width, height))
            using (var g = Graphics.FromImage(b))
            {
                g.Clear(Color.DodgerBlue);
                g.FillRectangle(Brushes.Red, 0, 0, width / 5, height);
                g.FillRectangle(Brushes.Lime, width * 4 / 5, 0, width / 5, height);
                b.Save(path, format);
            }
        }

        static List<Student> Cards(string photo)
        {
            return new List<Student>
            {
                new Student
                {
                    Record = 2,
                    DataRow = 1,
                    Name = "Fictional Student A",
                    Scholarship = "Fictional Test Scholarship & Award",
                    Message = "Dear Donor,\nThank you for supporting my education. Your generosity helps me purchase books and continue my academic journey.\nSincerely,\nFictional Student A",
                    Photo = photo
                },
                new Student
                {
                    Record = 3,
                    DataRow = 2,
                    Name = "Fictional Student B",
                    Scholarship = "Test Portfolio Two",
                    Message = String.Join(" ", Enumerable.Repeat("education", 310)) + " complete final sentence.",
                    Photo = ""
                }
            };
        }

        public static void Run(string root, string output)
        {
            Core.RejectLink(output);
            Directory.CreateDirectory(output);
            results = new List<string>();
            string template = Path.Combine(root, "Templates", "A7_Card_Template.docx"), hash = Core.Hash(template), photo = Path.Combine(output, "landscape.png");
            ImageFile(photo, 600, 300, ImageFormat.Png);
            Check("CSV quoted commas, embedded newlines and escaped quotes", () =>
            {
                var a = Core.ParseCsv("Name,Scholarship,Message,Photo\r\nA,\"Award, One\",\"Dear \"\"Donor\"\"\nThanks\",a.png\r\n");
                Assert(a.Count == 2 && a[1][1] == "Award, One" && a[1][2] == "Dear \"Donor\"\nThanks", "Quoted fields changed.");
            });
            Check("Malformed CSV quotes rejected", () =>
            {
                Reject(() => Core.ParseCsv("a,b\n\"bad"));
                Reject(() => Core.ParseCsv("\"a\"x,b"));
                Reject(() => Core.ParseCsv("a\"b,c"));
            });
            Check("Blank/repeated headings retained with unique labels", () =>
            {
                string p = Path.Combine(output, "headers.csv");
                File.WriteAllText(p, " Name ,NAME,,Message,Photo\nA,A,,Thanks,a.png\n,,,,\n", new UTF8Encoding(true));
                var d = Core.ReadCsv(p);
                Assert(d.Rows.Count == 1 && d.Columns.Count == 5 && d.Columns.Select(c => c.Label).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 5 && d.Columns[0].Status == "REPEATED_HEADER" && d.Columns[2].Status == "BLANK_HEADER", "Headings/blank rows wrong.");
            });
            Check("CSV column-count mismatch rejected", () =>
            {
                string p = Path.Combine(output, "wrong-width.csv");
                File.WriteAllText(p, "A,B,C,D\n1,2,3\n");
                Reject(() => Core.ReadCsv(p));
            });
            Check("Invalid UTF-8 and renamed Excel workbook rejected", () =>
            {
                string p = Path.Combine(output, "invalid.csv");
                File.WriteAllBytes(p, new byte[] { 255, 0, 128 });
                Reject(() => Core.ReadCsv(p));
                File.WriteAllBytes(p, new byte[] { 80, 75, 3, 4, 0 });
                Reject(() => Core.ReadCsv(p));
            });
            Check("Selecting two of 203 does not select all students", () =>
            {
                var s = Core.Select(203, 1, 2, false);
                Assert(s.Count == 2 && s.Start == 1 && s.End == 2, "Selected count changed.");
            });
            Check("Start/count, short final range and all-row selection", () =>
            {
                Assert(Core.Select(203, 51, 5, false).End == 55 && Core.Select(203, 202, 5, false).Count == 2 && Core.Select(203, 51, 5, true).Count == 203, "Range wrong.");
                Reject(() => Core.Select(203, 204, 1, false));
                Reject(() => Core.Select(203, 1, 0, false));
                Reject(() => Core.Select(0, 1, 1, false));
            });
            Check("Photo traversal, absolute paths and URLs rejected", () =>
            {
                foreach (string v in new[]
                {
                    "../outside.jpg",
                    "files/../../outside.jpg",
                    "C:\\outside.jpg",
                    "https://example.invalid/a.jpg",
                    "/outside.jpg",
                    "",
                    "a\u0000.jpg"
                }

                )
                    Reject(() => Core.PhotoPath(output, v));
                Assert(Core.PhotoPath(output, "files\\documents\\9001\\IMG_1231.jpg").StartsWith(Path.GetFullPath(output)), "Valid export path changed.");
            });
            Check("Unsupported photo documents and missing files rejected", () =>
            {
                Reject(() => Core.ResolvePhoto(output, "test.pdf"));
                Reject(() => Core.ResolvePhoto(output, "missing.jpg"));
            });
            Check("Broken photo bytes reported as decode failure", () =>
            {
                string p = Path.Combine(output, "broken.png");
                File.WriteAllText(p, "This is not an image.");
                Reject(() => Core.OpenPhoto(p));
            });
            Check("Supported standard image formats decode", () =>
            {
                foreach (var f in new[]
                {
                    Tuple.Create("jpg", ImageFormat.Jpeg),
                    Tuple.Create("png", ImageFormat.Png),
                    Tuple.Create("bmp", ImageFormat.Bmp),
                    Tuple.Create("gif", ImageFormat.Gif),
                    Tuple.Create("tif", ImageFormat.Tiff)
                }

                )
                {
                    string p = Path.Combine(output, "format." + f.Item1);
                    ImageFile(p, 80, 120, f.Item2);
                    using (var i = Core.OpenPhoto(p))
                        Assert(i.Image.Width == 80 && i.Image.Height == 120, "Decoded size wrong.");
                }
            });
            Check("WebP/HEIC/HEIF formats included with explicit codec failure", () =>
            {
                Assert(new[] { ".webp", ".heic", ".heif", ".jfif", ".jpe" }.All(e => Core.PhotoExtensions.Contains(e)), "Missing supported extension.");
                if (Environment.OSVersion.Platform != PlatformID.Win32NT)
                {
                    string p = Path.Combine(output, "codec.heic");
                    File.WriteAllBytes(p, new byte[4]);
                    try
                    {
                        Core.OpenPhoto(p);
                        throw new Exception("Invalid codec input accepted.");
                    }
                    catch (InvalidDataException e)
                    {
                        Assert(e.Message.StartsWith("CODEC_OR_PHOTO_ERROR"), "Codec error not clear.");
                    }
                }
            });
            Check("Original frame accepted; replacement or appended private metadata rejected", () =>
            {
                Core.ReadTemplate(template);
                byte[] frame = Entry(template, "word/media/image1.jpeg");
                byte[] note = Encoding.ASCII.GetBytes("Fictional private metadata");
                byte[] marked = new byte[frame.Length + note.Length + 4];
                marked[0] = 255;
                marked[1] = 216;
                marked[2] = 255;
                marked[3] = 254;
                marked[4] = (byte)((note.Length + 2) >> 8);
                marked[5] = (byte)(note.Length + 2);
                Buffer.BlockCopy(note, 0, marked, 6, note.Length);
                Buffer.BlockCopy(frame, 2, marked, 6 + note.Length, frame.Length - 2);
                string metadataFrame = Patch(template, Path.Combine(output, "frame-metadata.docx"), "word/media/image1.jpeg", marked);
                Reject(() => Core.ReadTemplate(metadataFrame));
                string replacementFrame = Patch(template, Path.Combine(output, "frame-replaced.docx"), "word/media/image1.jpeg", File.ReadAllBytes(photo));
                Reject(() => Core.ReadTemplate(replacementFrame));
            });
            var t = Core.ReadTemplate(template);
            var fitBox = Core.Box(t, Core.Modes[0], new Overlay());
            Check("Fit preserves entire photo and its aspect ratio", () =>
            {
                var p = Core.Plan(600, 300, fitBox, Core.Modes[0]);
                Assert(p.HorizontalCrop == 0 && p.VerticalCrop == 0 && Math.Abs(p.DisplayWidth / (double)p.DisplayHeight - 2) < .00001 && p.RenderWidth <= 600, "Fit resized/cropped incorrectly.");
            });
            Check("Fill uses centered crop within original opening", () =>
            {
                var p = Core.Plan(600, 300, fitBox, Core.Modes[1]);
                Assert(p.HorizontalCrop > 50 && p.SourceX > 0 && p.SourceY == 0 && p.DisplayWidth == (long)Math.Round(fitBox.Width) && p.DisplayHeight == (long)Math.Round(fitBox.Height), "Fill crop wrong.");
            });
            Check("Overlay permits bounded manual positions and rejects NaN/outside frame", () =>
            {
                var o = new Overlay
                {
                    Width = 2.1,
                    Height = 2.6,
                    X = .05,
                    Y = .05
                };
                var b = Core.Box(t, Core.Modes[2], o);
                Assert(b.X == .05 * 914400 && b.Width > fitBox.Width, "Overlay did not enlarge.");
                Reject(() => Core.Box(t, Core.Modes[2], new Overlay { Width = 10 }));
                Reject(() => Core.Box(t, Core.Modes[2], new Overlay { X = Double.NaN }));
                Reject(() => Core.Box(t, Core.Modes[2], new Overlay { X = .5 }));
            });
            Check("Preview and output share crop/geometry; source frame stays unchanged", () =>
            {
                foreach (string mode in Core.Modes)
                {
                    var p = Core.ReadPhoto(photo, Core.Box(t, mode, new Overlay()), mode);
                    using (var preview = Core.Preview(t, p))
                    {
                        Assert(preview.Width == 526 && preview.Height == 655, "Preview original frame size changed.");
                        preview.Save(Path.Combine(output, "Preview_" + Array.IndexOf(Core.Modes, mode) + ".png"), ImageFormat.Png);
                    }

                    Assert(p.Plan.RenderWidth <= 600 && p.Plan.RenderHeight <= 300, "Image was upsampled.");
                }

                Assert(Core.Hash(template) == hash, "Template changed.");
            });
            string merged = Path.Combine(output, "Merged_Test_Letters.docx");
            var cards = Cards(photo);
            Check("Two-card DOCX merge, one missing photo, all 310+ response words retained", () =>
            {
                Core.Build(Core.ReadTemplate(template), cards, merged, Core.Modes[2], new Overlay { X = .05, Y = .05 });
                Assert(cards.Count == 2 && cards[0].PhotoAdded && !cards[1].PhotoAdded && cards[1].Words > 300 && cards[1].Font == 10.5, "Merge/photo/font wrong.");
                Core.Validate(merged, cards);
                Assert(Core.ContentIssues(cards[1]).Any(s => s.StartsWith("WORD_LIMIT")) && Core.ContentIssues(cards[1]).Contains("PHOTO_MISSING"), "Content warnings missing.");
            });
            Check("Original template media, page size, fold and frame geometry preserved", () =>
            {
                Assert(Core.Hash(template) == hash, "Source bytes changed.");
                using (var original = ZipFile.OpenRead(template))
                    foreach (var e in original.Entries.Where(e => e.FullName.StartsWith("word/media/")))
                        Assert(Entry(template, e.FullName).SequenceEqual(Entry(merged, e.FullName)), "Original media changed.");
                using (var z = ZipFile.OpenRead(merged))
                {
                    var doc = Core.ReadXml(z, "word/document.xml");
                    var ns = Core.Names(doc);
                    var frames = doc.SelectNodes("//wp:anchor[@behindDoc='1' and .//a:blip]", ns);
                    Assert(frames.Count == 2, "Frame count wrong.");
                    foreach (XmlNode a in frames)
                    {
                        var e = (XmlElement)a.SelectSingleNode("wp:extent", ns);
                        Assert(e.GetAttribute("cx") == "2185200" && e.GetAttribute("cy") == "2718020", "Original frame resized.");
                    }

                    Assert(doc.SelectNodes("//wp:positionV/wp:posOffset[text()='4591988']", ns).Count == 2, "Fold line moved.");
                    var photos = doc.SelectNodes("//wp:anchor[@behindDoc='0' and .//a:blip]", ns);
                    Assert(photos.Count == 1 && ((XmlElement)photos[0]).GetAttribute("relativeHeight") == "251658241", "Photo not above original frame.");
                }
            });
            Check("Fit, fill and overlay merged packages all validate", () =>
            {
                foreach (string mode in Core.Modes)
                {
                    var list = Cards(photo);
                    Core.Build(Core.ReadTemplate(template), list, Path.Combine(output, "Mode_" + Array.IndexOf(Core.Modes, mode) + ".docx"), mode, new Overlay());
                    Assert(list[0].PhotoAdded, "Photo insertion failed: " + list[0].PhotoDetails);
                }
            });
            Check("Names and scholarship text mismatch rejected", () =>
            {
                var list = Cards(photo);
                list[0].Card = 1;
                list[1].Card = 2;
                list[0].Name = "Changed Name";
                Reject(() => Core.Validate(merged, list));
            });
            Check("Truncated student response rejected", () =>
            {
                var doc = new XmlDocument();
                doc.LoadXml(Encoding.UTF8.GetString(Entry(merged, "word/document.xml")));
                var ns = Core.Names(doc);
                var response = doc.SelectNodes("//w:t", ns).Cast<XmlNode>().First(n => n.InnerText.Contains("complete final sentence"));
                response.InnerText = "shortened";
                var p = Patch(merged, Path.Combine(output, "truncated.docx"), "word/document.xml", Core.XmlBytes(doc));
                Reject(() => Core.Validate(p, cards));
            });
            Check("Dangling image relationship rejected", () =>
            {
                var doc = new XmlDocument();
                doc.LoadXml(Encoding.UTF8.GetString(Entry(merged, "word/_rels/document.xml.rels")));
                var rel = doc.DocumentElement.ChildNodes.OfType<XmlElement>().First(e => e.GetAttribute("Type").EndsWith("/image"));
                rel.SetAttribute("Target", "media/missing-image.png");
                string p = Patch(merged, Path.Combine(output, "dangling.docx"), "word/_rels/document.xml.rels", Core.XmlBytes(doc));
                Reject(() => Core.Validate(p, cards));
            });
            Check("DTD/external entity template rejected", () =>
            {
                byte[] b = Encoding.UTF8.GetBytes("<!DOCTYPE document [<!ENTITY x SYSTEM 'file:///private'>]><document>&x;</document>");
                string p = Patch(template, Path.Combine(output, "dtd.docx"), "word/document.xml", b);
                Reject(() => Core.ReadTemplate(p));
            });
            Check("External resources in template rejected", () =>
            {
                var doc = new XmlDocument();
                doc.LoadXml(Encoding.UTF8.GetString(Entry(template, "word/_rels/document.xml.rels")));
                var rel = doc.DocumentElement.ChildNodes.OfType<XmlElement>().First();
                rel.SetAttribute("TargetMode", "External");
                rel.SetAttribute("Target", "https://example.invalid/payload");
                string p = Patch(template, Path.Combine(output, "external.docx"), "word/_rels/document.xml.rels", Core.XmlBytes(doc));
                Reject(() => Core.ReadTemplate(p));
            });
            Check("Macro/ActiveX/embedded objects and ZIP traversal rejected", () =>
            {
                int i = 0;
                foreach (string name in new[]
                {
                    "word/vbaProject.bin",
                    "word/activeX/activeX1.xml",
                    "word/embeddings/object.bin",
                    "word/media/payload.exe",
                    "../outside.xml"
                }

                )
                {
                    string p = Patch(template, Path.Combine(output, "unsafe_" + (i++) + ".docx"), name, new byte[] { 1 });
                    Reject(() => Core.ReadTemplate(p));
                }
            });
            Check("Active Word fields rejected", () =>
            {
                var doc = new XmlDocument();
                doc.LoadXml(Encoding.UTF8.GetString(Entry(template, "word/document.xml")));
                var ns = Core.Names(doc);
                doc.SelectSingleNode("//w:instrText", ns).InnerText = " DDEAUTO arbitrary command ";
                string p = Patch(template, Path.Combine(output, "active-field.docx"), "word/document.xml", Core.XmlBytes(doc));
                Reject(() => Core.ReadTemplate(p));
            });
            Check("Split active fields and fields outside main body rejected", () =>
            {
                var doc = new XmlDocument();
                doc.LoadXml(Encoding.UTF8.GetString(Entry(template, "word/document.xml")));
                var ns = Core.Names(doc);
                var paragraph = doc.CreateElement("w", "p", Core.W);
                foreach (string fragment in new[]
                {
                    " DDE",
                    "AUTO placeholder "
                }

                )
                {
                    var run = doc.CreateElement("w", "r", Core.W);
                    var instruction = doc.CreateElement("w", "instrText", Core.W);
                    instruction.InnerText = fragment;
                    run.AppendChild(instruction);
                    paragraph.AppendChild(run);
                }

                doc.SelectSingleNode("/w:document/w:body", ns).InsertBefore(paragraph, doc.SelectSingleNode("/w:document/w:body/w:sectPr", ns));
                string p = Patch(template, Path.Combine(output, "split-field.docx"), "word/document.xml", Core.XmlBytes(doc));
                Reject(() => Core.ReadTemplate(p));
                string header = "<w:hdr xmlns:w='" + Core.W + "'><w:p><w:r><w:instrText> DDEAUTO placeholder </w:instrText></w:r></w:p></w:hdr>";
                p = Patch(template, Path.Combine(output, "header-field.docx"), "word/header_test.xml", Encoding.UTF8.GetBytes(header));
                Reject(() => Core.ReadTemplate(p));
            });
            Check("Renamed macro content rejected by declared content type", () =>
            {
                var types = new XmlDocument();
                types.LoadXml(Encoding.UTF8.GetString(Entry(template, "[Content_Types].xml")));
                var entry = types.CreateElement("Override", types.DocumentElement.NamespaceURI);
                entry.SetAttribute("PartName", "/word/renamed.bin");
                entry.SetAttribute("ContentType", "application/vnd.ms-office.vbaProject");
                types.DocumentElement.AppendChild(entry);
                string p = Patch(template, Path.Combine(output, "renamed-macro.docx"), "[Content_Types].xml", Core.XmlBytes(types));
                Reject(() => Core.ReadTemplate(p));
            });
            Check("Repeated ZIP parts rejected", () =>
            {
                string p = Path.Combine(output, "duplicate.docx");
                using (var z = ZipFile.Open(p, ZipArchiveMode.Create))
                {
                    z.CreateEntry("a.xml");
                    z.CreateEntry("a.xml");
                }

                using (var z = ZipFile.OpenRead(p))
                    Reject(() => Core.CheckArchive(z));
            });
            Check("Spreadsheet formula injection escaped in reports", () =>
            {
                string p = Path.Combine(output, "report.csv");
                Core.Report(p, new[] { "Name", "Details" }, new[] { new[] { "=HYPERLINK(\"x\")", "normal, \"quoted\"" }, new[] { " +SUM(A1)", "message\nline" } });
                var rows = Core.ParseCsv(File.ReadAllText(p).TrimStart('\ufeff'));
                Assert(rows[1][0].StartsWith("'=") && rows[2][0].StartsWith("' +") && rows[1][1] == "normal, \"quoted\"" && rows[2][1] == "message\nline", "Report escaping failed.");
            });
            Check("Settings restore/rewrite only approved metadata", () =>
            {
                string p = Path.Combine(output, "settings.json");
                File.WriteAllText(p, "{\"Version\":1,\"StartRow\":51,\"RowCount\":2,\"PhotoMode\":\"Overlay curved corners\",\"Columns\":{\"Name\":\"Student name\"},\"Overlay\":{\"Width\":2.1,\"Height\":2.6,\"X\":0.05},\"StudentResponse\":\"private data should disappear\"}");
                var prefs = Settings.Read(p);
                Assert(prefs.StartRow == 51 && prefs.RowCount == 2 && prefs.Overlay.X == .05 && prefs.Columns["Name"] == "Student name", "Old preferences not restored.");
                Settings.Write(p, prefs);
                Settings.Write(p, prefs);
                Assert(!File.ReadAllText(p).Contains("private data"), "Unknown private field persisted.");
                Assert(Settings.Read(p).RowCount == 2, "Atomic rewrite failed.");
            });
            Check("Settings bounds and portable relative paths", () =>
            {
                string p = Path.Combine(output, "settings-invalid.json");
                File.WriteAllText(p, "{\"Version\":1,\"StartRow\":0,\"RowCount\":-1,\"PhotoMode\":\"invalid\",\"Overlay\":{\"X\":9}}");
                var prefs = Settings.Read(p);
                Assert(prefs.StartRow == 1 && prefs.RowCount == 5 && prefs.PhotoMode == Core.Modes[0] && prefs.Overlay.X == 0, "Invalid preferences accepted.");
                string path = Settings.StorePath(output, photo);
                Assert(!Path.IsPathRooted(path) && Settings.RestorePath(output, path, ".png") == Path.GetFullPath(photo), "Portable path wrong.");
            });
            Check("Managed executable has no embedded PowerShell or native P/Invoke methods", () =>
            {
                var a = typeof(Core).Assembly;
                Assert(!a.GetReferencedAssemblies().Any(r => r.Name.Contains("Management.Automation")), "PowerShell assembly reference found.");
                foreach (var type in a.GetTypes())
                    foreach (var m in type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance))
                        Assert((m.Attributes & MethodAttributes.PinvokeImpl) == 0, "Native imported method found.");
            });
            Check("Word layout availability reported honestly", () =>
            {
                if (Environment.OSVersion.Platform != PlatformID.Win32NT)
                {
                    var report = WordLayout.Check(merged, cards, t.Fold, true, output, s =>
                    {
                    });
                    Assert(report.Count == 2 && report.All(r => r.Status == "LAYOUT_UNCHECKED") && !File.Exists(Path.Combine(output, "Merged_Letters.pdf")), "Unavailable Word reported as checked.");
                }
            });
            Check("Production assembly excludes test fixtures", () =>
            {
                Assert(typeof(Core).Assembly.GetType("LetterMerger.Tests") == null, "Test fixtures bundled in production executable.");
            });
            Check("Fit zoom centers both portrait and landscape and adapts to resize", () =>
            {
                var view = new PreviewViewState();
                view.SetImageSize(new Size(1000, 500));
                view.SetViewport(new Size(800, 600));
                Assert(Math.Abs(view.Zoom - .8) < .00001 && view.Bounds.X == 0 && view.Bounds.Y == 100, "Landscape fit incorrect.");
                view.SetViewport(new Size(400, 300));
                Assert(Math.Abs(view.Zoom - .4) < .00001, "Fit did not adapt to resize.");
                view.SetImageSize(new Size(500, 1000));
                Assert(Math.Abs(view.Zoom - .3) < .00001 && view.Bounds.X == 125, "Portrait fit incorrect.");
            });
            Check("Manual zoom limits, invalid values and minimize-sized viewport", () =>
            {
                var view = new PreviewViewState();
                view.SetImageSize(new Size(1000, 1000));
                view.SetViewport(new Size(400, 400));
                view.SetZoom(.001);
                Assert(view.Zoom == .1, "Minimum zoom incorrect.");
                view.SetZoom(999);
                Assert(view.Zoom == 8, "Maximum zoom incorrect.");
                Reject(() => view.SetZoom(Double.NaN));
                Reject(() => view.SetZoom(Double.PositiveInfinity));
                Reject(() => view.SetZoom(0));
                view.SetZoom(2);
                view.SetViewport(Size.Empty);
                view.SetViewport(new Size(800, 600));
                Assert(view.Zoom == 2 && !view.FitMode, "Manual zoom lost after restore.");
                view.SetImageSize(new Size(800, 1200));
                Assert(view.Zoom == 2 && view.PanX == 0 && view.PanY == 0, "New photo should retain zoom and recenter.");
            });
            Check("Wheel zoom anchors the photo under the pointer", () =>
            {
                var view = new PreviewViewState();
                view.SetImageSize(new Size(1000, 1000));
                view.SetViewport(new Size(400, 400));
                view.SetZoom(2);
                var before = view.Bounds;
                var pointer = new PointF(150, 100);
                double x = (pointer.X - before.X) / view.Zoom, y = (pointer.Y - before.Y) / view.Zoom;
                view.SetZoomAt(4, pointer);
                Assert(Math.Abs((pointer.X - view.Bounds.X) / view.Zoom - x) < .001 && Math.Abs((pointer.Y - view.Bounds.Y) / view.Zoom - y) < .001, "Pointer anchor drifted.");
            });
            Check("Pan stays bounded and Fit restores centered view", () =>
            {
                var view = new PreviewViewState();
                view.SetImageSize(new Size(1000, 500));
                view.SetViewport(new Size(400, 400));
                view.SetZoom(1);
                view.Pan(100000, -100000);
                Assert(view.PanX == 300 && view.PanY == -50 && view.Bounds.Right >= 400 && view.Bounds.Top <= 0, "Pan can lose the image.");
                view.SetZoom(.2);
                Assert(view.PanX == 0 && view.PanY == 0, "Small image should center.");
                view.Fit();
                view.Pan(10, 10);
                Assert(view.FitMode && view.PanX == 0 && view.PanY == 0, "Fit did not reset panning.");
            });
            Check("Common punctuation and accents repair with exact full-response preservation", () =>
            {
                string expected = "Dear donor,\nI\u2019m grateful \u2013 caf\u00e9 studies help me. \u201cThank you\u201d for \u20ac100.\nWith gratitude";
                string broken = Encoding.GetEncoding(1252).GetString(Encoding.UTF8.GetBytes(expected));
                var fixedText = TextRepair.Fix(broken);
                Assert(fixedText.Text == expected && fixedText.Changes.Count > 0 && !fixedText.Unresolved, "Common encoding repair lost text.");
                string doubled = Encoding.GetEncoding(1252).GetString(Encoding.UTF8.GetBytes(broken));
                Assert(TextRepair.Fix(doubled).Text == expected, "Double-encoded text not repaired.");
                string latin = Encoding.GetEncoding(28591).GetString(Encoding.UTF8.GetBytes(expected));
                Assert(TextRepair.Fix(latin).Text == expected, "Latin-1 mojibake not repaired.");
            });
            Check("Mixed multilingual text and valid Unicode remain intact", () =>
            {
                string correct = "Jos\u00e9, Fran\u00e7ois, \u00c2ngela, \u00c3 as a symbol, \u4e2d\u6587, \ud83d\ude42, \u2013 and \u2019. French: \u00e9\u00a0\u00bb.";
                Assert(TextRepair.Fix(correct).Text == correct && TextRepair.Fix(correct).Changes.Count == 0, "Valid Unicode changed.");
                string mixed = correct + " Support \u00e2\u20ac\u201c education: caf\u00c3\u00a9.";
                Assert(TextRepair.Fix(mixed).Text == correct + " Support \u2013 education: caf\u00e9.", "Mixed Unicode damaged.");
            });
            Check("Irrecoverable and incomplete characters are flagged without guessing", () =>
            {
                foreach (string text in new[] { "Lost \ufffd letter", "Incomplete \u00e2\u20ac sequence" })
                {
                    var repaired = TextRepair.Fix(text);
                    Assert(repaired.Text == text && repaired.Unresolved, "Incomplete sequence was changed or hidden.");
                }
            });
            Check("Repair option off and exact name/scholarship protections", () =>
            {
                var student = Cards(photo)[0];
                student.Name = "Jos\u00c3\u00a9";
                student.Scholarship = "Caf\u00c3\u00a9 Award";
                student.Message = "Thank you \u00e2\u20ac\u201c I appreciate it.";
                string original = student.Message;
                TextRepair.Prepare(student, false);
                Assert(student.Message == original && student.Name == "Jos\u00c3\u00a9" && student.Scholarship == "Caf\u00c3\u00a9 Award" && student.TextIssues.Count == 3, "Repair disabled/identity protections failed.");
                Assert(TextRepair.ReportRows(new[] { student }).Any(r => r[5] == "REPAIR_DISABLED"), "Disabled repairs not reported.");
            });
            Check("Repaired merged DOCX keeps full message and rejects unrelated edits", () =>
            {
                var students = Cards(photo);
                students[0].Message = "Dear donor,\nI\u00e2\u20ac\u2122m grateful \u00e2\u20ac\u201c thank you.\nSincerely,\nFictional Student A";
                TextRepair.Prepare(students[0], true);
                TextRepair.Prepare(students[1], true);
                string path = Path.Combine(output, "repaired.docx");
                Core.Build(Core.ReadTemplate(template), students, path, Core.Modes[0], new Overlay());
                Core.Validate(path, students);
                Assert(Encoding.UTF8.GetString(Entry(path, "word/document.xml")).Contains("I\u2019m grateful \u2013 thank you."), "DOCX still contains broken punctuation.");
                students[0].Message += " unauthorized rewrite";
                Reject(() => Core.Build(Core.ReadTemplate(template), students, Path.Combine(output, "unapproved.docx"), Core.Modes[0], new Overlay()));
                Assert(!File.Exists(Path.Combine(output, "unapproved.docx")), "Unapproved response edit published.");
            });
            Check("Split photo/missing DOCX groups preserve local card indices and validation", () =>
            {
                string folder = Path.Combine(output, "split-mixed"); Directory.CreateDirectory(folder);
                var students = Cards(photo);
                var batch = BatchMerge.Build(template, students, folder, Core.Modes[0], new Overlay(), true, message => { });
                Assert(batch.Errors.Count == 0 && batch.Outputs.Count == 2 && batch.Outputs.Sum(o => o.Cards.Count) == 2, "Mixed split failed.");
                Assert(students[0].OutputFile.EndsWith(BatchMerge.MainFile) && students[0].PhotoAdded && students[1].OutputFile.EndsWith(BatchMerge.MissingFile) && !students[1].PhotoAdded && students.All(c => c.Card == 1), "Split assignment/card numbering wrong.");
                foreach (var document in batch.Outputs)
                {
                    Core.Validate(document.Path, document.Cards);
                    var rows = WordLayout.Check(document.Path, document.Cards, document.Fold, true, folder, message => { });
                    Assert(rows.All(r => r.OutputFile == document.Path), "Layout report lacks document association.");
                }
                Assert(File.Exists(Path.Combine(folder, "DocumentValidation_Merged_Letters.txt")) && File.Exists(Path.Combine(folder, "DocumentValidation_Merged_Letters_Missing.txt")) && !Directory.GetDirectories(folder).Any(), "Validation collision or temp directory leaked.");
            });
            Check("All missing, all photos, empty batch and separation off produce no empty DOCX", () =>
            {
                foreach (string kind in new[] { "missing", "photos", "combined", "empty" })
                {
                    string folder = Path.Combine(output, "group-" + kind); Directory.CreateDirectory(folder);
                    var students = Cards(photo);
                    if (kind == "missing") students[0].Photo = "";
                    if (kind == "photos") students[1].Photo = photo;
                    if (kind == "empty") students.Clear();
                    var batch = BatchMerge.Build(template, students, folder, Core.Modes[1], new Overlay(), kind != "combined", message => { });
                    Assert(batch.Errors.Count == 0 && batch.Outputs.Count == (kind == "empty" ? 0 : 1), "Empty/unified group count wrong.");
                    string expected = kind == "missing" ? BatchMerge.MissingFile : BatchMerge.MainFile;
                    if (kind != "empty") Assert(batch.Outputs[0].Path.EndsWith(expected) && batch.Outputs[0].Cards.Count == 2, "Empty file or wrong group emitted.");
                    Assert(Directory.GetFiles(folder, "*.docx").Length == (kind == "empty" ? 0 : 1), "Unnecessary DOCX produced.");
                }
            });
            Check("Late photo insertion failure is regrouped into missing file without loss", () =>
            {
                string folder = Path.Combine(output, "late-photo-failure"); Directory.CreateDirectory(folder);
                var students = Cards(photo);
                var late = new Student { Record = 4, DataRow = 3, Name = "Late Failure Student", Scholarship = "Test Award", Message = "Thank you for supporting my education.", Photo = Path.Combine(output, "late-broken.jpg") };
                File.WriteAllText(late.Photo, "broken image bytes");
                students.Add(late);
                var batch = BatchMerge.Build(template, students, folder, Core.Modes[2], new Overlay(), true, message => { });
                Assert(batch.Errors.Count == 0 && batch.Outputs.Count == 2 && batch.Outputs.Sum(o => o.Cards.Count) == 3, "Late failure lost a card.");
                Assert(late.PhotoStatus == "PHOTO_INSERT_ERROR" && late.OutputFile.EndsWith(BatchMerge.MissingFile), "Late failure remained in photo-complete output.");
                Assert(batch.Outputs.Single(o => o.Path.EndsWith(BatchMerge.MainFile)).Cards.All(c => c.PhotoAdded), "Main file has missing-photo card.");
                foreach (var document in batch.Outputs) Core.Validate(document.Path, document.Cards);
            });
            Check("Partial output failure preserves successful group and truthful student file paths", () =>
            {
                string folder = Path.Combine(output, "partial-failure"); Directory.CreateDirectory(folder);
                File.WriteAllText(Path.Combine(folder, BatchMerge.MainFile), "existing file must not be overwritten");
                var students = Cards(photo);
                var batch = BatchMerge.Build(template, students, folder, Core.Modes[0], new Overlay(), true, message => { });
                Assert(batch.Errors.Count > 0 && batch.Outputs.Count == 1 && students[0].OutputFile == "" && students[1].OutputFile.EndsWith(BatchMerge.MissingFile), "Partial failure reported as complete.");
                Assert(File.ReadAllText(Path.Combine(folder, BatchMerge.MainFile)) == "existing file must not be overwritten", "Existing output overwritten.");
            });
            Check("New options migrate from older settings and persist with Restore Defaults values", () =>
            {
                string path = Path.Combine(output, "settings-v14.json");
                File.WriteAllText(path, "{\"Version\":1,\"RowCount\":2}");
                var settings = Settings.Read(path);
                Assert(settings.RepairText && !settings.SeparateMissingPhotos && settings.RowCount == 2, "Old settings did not migrate.");
                settings.RepairText = false; settings.SeparateMissingPhotos = true;
                Settings.Write(path, settings);
                settings = Settings.Read(path);
                Assert(!settings.RepairText && settings.SeparateMissingPhotos, "New options did not persist.");
            });
            Check("Template comments, custom XML, attachments and additional media are rejected", () =>
            {
                int number = 0;
                foreach (string part in new[] { "word/comments.xml", "customXml/item1.xml", "docProps/custom.xml", "word/media/unused.jpg", "word/header1.xml", "private-record.txt" })
                {
                    string path = Patch(template, Path.Combine(output, "private-extra-" + number++ + ".docx"), part, Encoding.UTF8.GetBytes("fictional sensitive fixture"));
                    Reject(() => Core.ReadTemplate(path));
                }
            });
            Check("Filled-in template text and tracked or hidden text are rejected", () =>
            {
                foreach (string kind in new[] { "filled", "hidden", "tracked" })
                {
                    var document = new XmlDocument();
                    document.LoadXml(Encoding.UTF8.GetString(Entry(template, "word/document.xml")));
                    var ns = Core.Names(document);
                    if (kind == "filled") document.SelectSingleNode("//w:t", ns).InnerText = "Fictional private student";
                    else if (kind == "hidden")
                    {
                        var run = document.SelectSingleNode("//w:r", ns);
                        Core.Property(Core.Property(run, "rPr"), "vanish");
                    }
                    else document.SelectSingleNode("/w:document/w:body", ns).AppendChild(document.CreateElement("w", "ins", Core.W));
                    string path = Patch(template, Path.Combine(output, "private-content-" + kind + ".docx"), "word/document.xml", Core.XmlBytes(document));
                    Reject(() => Core.ReadTemplate(path));
                }
            });
            Check("Template footnotes cannot carry hidden records", () =>
            {
                var document = new XmlDocument();
                document.LoadXml(Encoding.UTF8.GetString(Entry(template, "word/footnotes.xml")));
                var ns = Core.Names(document);
                var text = document.CreateElement("w", "t", Core.W);
                text.InnerText = "Fictional private note";
                document.SelectSingleNode("//w:r", ns).AppendChild(text);
                string path = Patch(template, Path.Combine(output, "private-note.docx"), "word/footnotes.xml", Core.XmlBytes(document));
                Reject(() => Core.ReadTemplate(path));
            });
            Check("Mixed-case external resources and malformed internal targets are rejected before preview", () =>
            {
                foreach (string kind in new[] { "external", "missing", "scheme" })
                {
                    var document = new XmlDocument();
                    document.LoadXml(Encoding.UTF8.GetString(Entry(template, "word/_rels/document.xml.rels")));
                    var relation = document.DocumentElement.ChildNodes.OfType<XmlElement>().First();
                    relation.SetAttribute("Target", kind == "external" ? "https://example.invalid/resource" : kind == "scheme" ? "file:///unrelated" : "media/missing.png");
                    if (kind == "external") relation.SetAttribute("TargetMode", "eXtErNaL");
                    string path = Patch(template, Path.Combine(output, "unsafe-target-" + kind + ".docx"), "word/_rels/document.xml.rels", Core.XmlBytes(document));
                    Reject(() => Core.ReadTemplate(path));
                }
            });
            Check("Case-conflicting archive entries are rejected", () =>
            {
                string path = Patch(template, Path.Combine(output, "case-conflict.docx"), "word/Document.xml", Entry(template, "word/document.xml"));
                Reject(() => Core.ReadTemplate(path));
            });
            Check("Merged output clears inherited document author and history metadata", () =>
            {
                var document = new XmlDocument();
                document.LoadXml(Encoding.UTF8.GetString(Entry(template, "docProps/core.xml")));
                var creator = document.CreateElement("dc", "creator", "http://purl.org/dc/elements/1.1/");
                creator.InnerText = "Fictional private author";
                document.DocumentElement.AppendChild(creator);
                string source = Patch(template, Path.Combine(output, "private-author.docx"), "docProps/core.xml", Core.XmlBytes(document));
                string target = Path.Combine(output, "clean-metadata.docx");
                Core.Build(Core.ReadTemplate(source), Cards(photo), target, Core.Modes[0], new Overlay());
                string metadata = Encoding.UTF8.GetString(Entry(target, "docProps/core.xml")) + Encoding.UTF8.GetString(Entry(target, "docProps/app.xml"));
                Assert(!metadata.Contains("creator") && !metadata.Contains("lastModifiedBy") && !metadata.Contains("Fictional private author") && !metadata.Contains("created"), "Template metadata leaked.");
            });
            Check("Package writing never overwrites an existing destination", () =>
            {
                string path = Path.Combine(output, "existing-package.docx");
                File.WriteAllText(path, "keep this existing file");
                Reject(() => Core.WritePackage(template, path, new Dictionary<string, byte[]>()));
                Assert(File.ReadAllText(path) == "keep this existing file", "Destination overwritten.");
            });
            Check("Formula-like CSV column headings and cells are escaped", () =>
            {
                string path = Path.Combine(output, "safe-headings.csv");
                Core.Report(path, new[] { "=formula", "Safe" }, new[] { new[] { "\t=payload", "normal", "\tcmd", "\rcmd", " +1", "plain" } });
                var rows = Core.ParseCsv(File.ReadAllText(path).TrimStart('\ufeff'));
                Assert(rows[0][0].StartsWith("'=") && rows[1][0].StartsWith("'\t="), "CSV header or tab-prefix protection failed.");
                Assert(rows[1][2] == "'\tcmd" && rows[1][3] == "'\rcmd" && rows[1][4] == "' +1" && rows[1][5] == "plain", "Leading tab, carriage return, or spaced formula protection failed.");
            });
            Check("An existing partial output is preserved on merge failure", () =>
            {
                string path = Path.Combine(output, "preserved-partial.docx");
                File.WriteAllText(path + ".partial", "preserve this earlier partial output");
                Reject(() => Core.Build(Core.ReadTemplate(template), Cards(photo), path, Core.Modes[0], new Overlay()));
                Assert(File.ReadAllText(path + ".partial") == "preserve this earlier partial output" && !File.Exists(path), "Earlier partial output was deleted or published.");
            });
            Check("Non-raster files disguised as JPG are rejected before native decoding", () =>
            {
                string path = Path.Combine(output, "disguised-vector.jpg");
                File.WriteAllBytes(path, new byte[] { 215, 205, 198, 154, 0, 0, 0, 0 });
                try { Core.OpenPhoto(path); throw new InvalidOperationException("Vector content accepted."); }
                catch (InvalidDataException error) { Assert(error.Message.StartsWith("UNSUPPORTED_PHOTO_CONTENT"), "Vector content reached the native decoder."); }
            });
            Check("Different names retain photo ownership distinctions", () =>
            {
                Assert(Core.PhotoOwner("Fictional A-B") != Core.PhotoOwner("Fictional AB"), "Punctuation stripped from identity.");
                Assert(Core.PhotoOwner(" A ") == "A" && Core.PhotoOwner("Jos\u00e9") == Core.PhotoOwner("Jose\u0301"), "Identity normalization failed.");
            });
            Check("Merged-batch limits allow many photos without weakening template limits", () =>
            {
                string path = Path.Combine(output, "large-batch-size.docx");
                var bytes = new byte[1024 * 1024];
                using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
                    for (int i = 0; i < 160; i++)
                        using (var stream = archive.CreateEntry("media/fictional-" + i + ".png", CompressionLevel.Optimal).Open())
                            stream.Write(bytes, 0, bytes.Length);
                using (var archive = ZipFile.OpenRead(path))
                {
                    Reject(() => Core.CheckArchive(archive));
                    Core.CheckArchive(archive, true);
                }
            });
            results.Add("LIMIT: Windows Defender, desktop Word pagination and installed HEIC/HEIF/WebP codecs require testing on Windows. Core self-checks do not certify antivirus safety.");
            File.WriteAllLines(Path.Combine(output, "SelfCheckResults.txt"), results, new UTF8Encoding(true));
        }

        static IEnumerable<Control> Descendants(Control c)
        {
            foreach (Control child in c.Controls)
            {
                yield return child;
                foreach (var d in Descendants(child))
                    yield return d;
            }
        }

        static void Snapshot(Form form, string path)
        {
            Application.DoEvents();
            form.Location = new Point(10, 10);
            form.BringToFront();
            form.Refresh();
            Application.DoEvents();
            File.WriteAllLines(path + ".layout.txt", Descendants(form).Select(c => c.GetType().Name + " " + c.Text.Replace("\r", " ").Replace("\n", " ") + " " + c.Bounds + " Visible=" + c.Visible));
            var surface = form.Controls.OfType<TableLayoutPanel>().Single();
            using (var bitmap = new Bitmap(surface.ClientSize.Width, surface.ClientSize.Height))
            {
                surface.DrawToBitmap(bitmap, surface.ClientRectangle);
                bitmap.Save(path, ImageFormat.Png);
            }
        }

        public static void UI(string root, string output)
        {
            Directory.CreateDirectory(output);
            string template = Path.Combine(output, "Templates", "A7_Card_Template.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(template));
            File.Copy(Path.Combine(root, "Templates", "A7_Card_Template.docx"), template, false);
            Directory.CreateDirectory(Path.Combine(output, "Data"));
            string csv = Path.Combine(output, "Data", "test.csv");
            File.WriteAllText(csv, "Name,Portfolio Name,Message,Photo Path\n Fictional Student , Test Scholarship ,Thank you \u00e2\u20ac\u201c for your gift.,preview.jpg\nAnother Student,Test Scholarship,Thank you.,missing.jpg\n");
            string sentinel = Path.Combine(output, "Photos", "keep.txt");
            using (var form = new MainForm(output))
            {
                File.WriteAllText(sentinel, "preserve this file");
                ImageFile(Path.Combine(output, "Photos", "preview.jpg"), 400, 600, ImageFormat.Jpeg);
                form.Show();
                Application.DoEvents();
                Assert(form.Text == "Letter Merger", "Title wrong.");
                var controls = Descendants(form).ToList();
                Assert(controls.OfType<Label>().Any(l => l.Text == "-Kevin Le"), "Watermark missing.");
                var csvBox = controls.OfType<ComboBox>().First(c => c.DisplayMember == "Name" && c.Items.Cast<object>().Any(i => i is FileInfo && ((FileInfo)i).Name == "test.csv"));
                Assert(form.MaximizeBox && form.MinimizeBox && form.FormBorderStyle == FormBorderStyle.Sizable, "Main window cannot resize/minimize/maximize.");
                form.ClientSize = new Size(800, 560);
                Application.DoEvents();
                int smallWidth = csvBox.Width;
                var mainShell = form.Controls.OfType<TableLayoutPanel>().Single();
                int smallClientWidth = mainShell.ClientSize.Width;
                Snapshot(form, Path.Combine(output, "Main_Compact.png"));
                form.ClientSize = new Size(1500, 950);
                Application.DoEvents();
                Assert(mainShell.ClientSize.Width > smallClientWidth + 100 && csvBox.Width - smallWidth >= (mainShell.ClientSize.Width - smallClientWidth) * .75, "Main file controls did not grow proportionally. Available client: " + smallClientWidth + " -> " + mainShell.ClientSize.Width + "; CSV: " + smallWidth + " -> " + csvBox.Width + "; Screen: " + Screen.PrimaryScreen.WorkingArea);
                Assert(controls.OfType<ComboBox>().All(c => c.DropDownWidth <= c.Width), "Dropdown extends past control width.");
                var footer = controls.OfType<Label>().First(l => l.Text == "-Kevin Le");
                var point = form.PointToClient(footer.PointToScreen(Point.Empty));
                Assert(point.Y > mainShell.ClientSize.Height - 35, "Footer detached from bottom.");
                typeof(MainForm).GetMethod("SetBusy", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { true });
                Assert(form.ControlBox && form.MinimizeBox && form.MaximizeBox && !csvBox.Enabled, "Busy state disabled window chrome or left inputs active.");
                typeof(MainForm).GetMethod("SetBusy", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { false });
                Assert(csvBox.Enabled, "Inputs not restored.");
                form.WindowState = FormWindowState.Minimized;
                Application.DoEvents();
                form.WindowState = FormWindowState.Normal;
                Application.DoEvents();
                Snapshot(form, Path.Combine(output, "Main_Large.png"));
                form.RestoreDefaults();
                Assert(form.DefaultsCorrect() && File.ReadAllText(sentinel) == "preserve this file", "Restore Defaults wrong/deleted input.");
                typeof(MainForm).GetMethod("LoadColumns", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { true });
                string csvHash = Core.Hash(csv);
                var repairOption = controls.OfType<CheckBox>().Single(c => c.Text.StartsWith("Repair common"));
                var separateOption = controls.OfType<CheckBox>().Single(c => c.Text.StartsWith("Put letters"));
                Assert(repairOption.Checked && !separateOption.Checked, "New default options wrong.");
                separateOption.Checked = true;
                typeof(MainForm).GetMethod("Run", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { true });
                string generatedBatch = Directory.GetDirectories(Path.Combine(output, "Output"), "Batch_*").Single();
                Assert(File.Exists(Path.Combine(generatedBatch, BatchMerge.MainFile)) && File.Exists(Path.Combine(generatedBatch, BatchMerge.MissingFile)), "UI merge did not split output.");
                var mergeRows = Core.ParseCsv(File.ReadAllText(Path.Combine(generatedBatch, "MergeReport.csv")).TrimStart('\ufeff'));
                int fileColumn = Array.IndexOf(mergeRows[0], "File"), cardColumn = Array.IndexOf(mergeRows[0], "Card");
                Assert(mergeRows.Count == 3 && mergeRows[1][fileColumn].EndsWith(BatchMerge.MainFile) && mergeRows[2][fileColumn].EndsWith(BatchMerge.MissingFile) && mergeRows[1][cardColumn] == "1" && mergeRows[2][cardColumn] == "1", "UI merge report has wrong counts/file/card.");
                Assert(mergeRows[1][Array.IndexOf(mergeRows[0], "StudentName")] == " Fictional Student " && mergeRows[1][Array.IndexOf(mergeRows[0], "Scholarship")] == " Test Scholarship ", "Name or scholarship whitespace was altered.");
                var layoutRows = Core.ParseCsv(File.ReadAllText(Path.Combine(generatedBatch, "LayoutReport.csv")).TrimStart('\ufeff'));
                Assert(layoutRows.Count == 3 && layoutRows[0].Last() == "OutputFile" && layoutRows[1].Last().EndsWith(BatchMerge.MainFile) && layoutRows[2].Last().EndsWith(BatchMerge.MissingFile), "Layout report associations wrong.");
                var repairRows = Core.ParseCsv(File.ReadAllText(Path.Combine(generatedBatch, "TextRepairReport.csv")).TrimStart('\ufeff'));
                Assert(repairRows.Count == 2 && repairRows[1][5] == "REPAIRED" && Core.Hash(csv) == csvHash, "Repairs were not reported or modified CSV.");
                Assert(Settings.Read(Path.Combine(output, "LetterMergerSettings.json")).SeparateMissingPhotos, "UI option not saved.");
                form.RestoreDefaults();
                Assert(form.DefaultsCorrect(), "Restore Defaults did not reset new options.");
                var data = Core.ReadCsv(csv);
                var map = new Dictionary<string, int> { { "Name", 0 }, { "Scholarship", 1 }, { "Message", 2 }, { "Photo", 3 } };
                var overlay = new Overlay();
                bool applied = false;
                using (var preview = new PhotoPreview(Core.ReadTemplate(template), data, map, Core.Select(2, 1, 2, false), Path.Combine(output, "Photos"), overlay, Core.Modes[0], choice => applied = true))
                {
                    preview.Show();
                    Application.DoEvents();
                    Assert(preview.MinimizeBox && preview.MaximizeBox && preview.ShowInTaskbar && preview.FormBorderStyle == FormBorderStyle.Sizable, "Preview window controls unavailable.");
                    var view = Descendants(preview).OfType<ZoomPhotoView>().Single();
                    Assert(view.Image != null, "Photo preview failed.");
                    preview.ClientSize = new Size(800, 510);
                    Application.DoEvents();
                    var small = view.Size;
                    var previewShell = preview.Controls.OfType<TableLayoutPanel>().Single();
                    var smallClient = previewShell.ClientSize;
                    Snapshot(preview, Path.Combine(output, "Preview_Compact.png"));
                    preview.ClientSize = new Size(1450, 950);
                    Application.DoEvents();
                    Assert(previewShell.ClientSize.Width > smallClient.Width + 100 && previewShell.ClientSize.Height > smallClient.Height + 100 && view.Width - small.Width >= (previewShell.ClientSize.Width - smallClient.Width) * .75 && view.Height - small.Height >= (previewShell.ClientSize.Height - smallClient.Height) * .75, "Preview canvas did not grow proportionally in both dimensions. Available client: " + smallClient + " -> " + previewShell.ClientSize + "; Canvas: " + small + " -> " + view.Size);
                    Snapshot(preview, Path.Combine(output, "Preview_Large.png"));
                    var previewControls = Descendants(preview).ToList();
                    var slider = previewControls.OfType<TrackBar>().Single();
                    var wheel = typeof(ZoomPhotoView).GetMethod("OnMouseWheel", BindingFlags.NonPublic | BindingFlags.Instance);
                    wheel.Invoke(view, new object[] { new MouseEventArgs(MouseButtons.None, 0, view.Width / 2, view.Height / 2, 120) });
                    Assert(!view.State.FitMode, "Mouse wheel did not zoom.");
                    slider.Value = 200;
                    Application.DoEvents();
                    Assert(view.State.Zoom == 2 && !view.State.FitMode, "Slider zoom failed.");
                    double originalPan = view.State.PanX, originalPanY = view.State.PanY;
                    typeof(ZoomPhotoView).GetMethod("OnMouseDown", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(view, new object[] { new MouseEventArgs(MouseButtons.Left, 1, 100, 100, 0) });
                    typeof(ZoomPhotoView).GetMethod("OnMouseMove", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(view, new object[] { new MouseEventArgs(MouseButtons.Left, 0, 120, 110, 0) });
                    typeof(ZoomPhotoView).GetMethod("OnMouseUp", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(view, new object[] { new MouseEventArgs(MouseButtons.Left, 1, 120, 110, 0) });
                    Assert(view.State.PanX != originalPan || view.State.PanY != originalPanY, "Drag did not pan the zoomed view.");
                    previewControls.OfType<Button>().Single(b => b.Text == "Zoom +").PerformClick();
                    Assert(Math.Abs(view.State.Zoom - 2.4) < .001, "Zoom + failed.");
                    previewControls.OfType<Button>().Single(b => b.Text == "Zoom -").PerformClick();
                    Assert(Math.Abs(view.State.Zoom - 2) < .001, "Zoom - failed.");
                    preview.WindowState = FormWindowState.Minimized;
                    Application.DoEvents();
                    preview.WindowState = FormWindowState.Normal;
                    preview.ClientSize = new Size(1250, 850);
                    Application.DoEvents();
                    Assert(view.State.Zoom == 2, "Manual zoom lost on resize/restore.");
                    Snapshot(preview, Path.Combine(output, "Preview_Zoomed.png"));
                    previewControls.OfType<Button>().Single(b => b.Text == "Fit").PerformClick();
                    Assert(view.State.FitMode, "Fit button failed.");
                    previewControls.OfType<Button>().Single(b => b.Text == "100%").PerformClick();
                    Assert(view.State.Zoom == 1, "100% button failed.");
                    previewControls.OfType<Button>().Single(b => b.Text == "Next").PerformClick();
                    Assert(view.Image == null && previewControls.OfType<Label>().Any(l => l.Text.StartsWith("Preview issue:")), "Missing row left stale photo.");
                    previewControls.OfType<Button>().Single(b => b.Text == "Previous").PerformClick();
                    Assert(view.Image != null, "Photo did not return after missing row.");
                    Assert(!applied && overlay.Width == 2.10 && overlay.Height == 2.60 && overlay.X == 0 && overlay.Y == 0, "View zoom modified batch geometry.");
                    preview.Close();
                }
                typeof(MainForm).GetMethod("LoadColumns", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { true });
                var openPreview = typeof(MainForm).GetMethod("Preview", BindingFlags.NonPublic | BindingFlags.Instance);
                openPreview.Invoke(form, null);
                Application.DoEvents();
                var active = (PhotoPreview)typeof(MainForm).GetField("activePreview", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
                Assert(active != null && active.Visible && active.Owner == null && !csvBox.Enabled && form.ControlBox, "Main preview did not open as an independent modeless window.");
                form.WindowState = FormWindowState.Minimized;
                Application.DoEvents();
                Assert(active.WindowState != FormWindowState.Minimized, "Main minimize forced preview minimize.");
                form.WindowState = FormWindowState.Normal;
                openPreview.Invoke(form, null);
                Application.DoEvents();
                Assert(Application.OpenForms.Cast<Form>().Count(f => f is PhotoPreview) == 1, "Duplicate preview window opened.");
                active.Close();
                Application.DoEvents();
                Assert(csvBox.Enabled, "Main controls stayed disabled after preview close.");
                openPreview.Invoke(form, null);
                active = (PhotoPreview)typeof(MainForm).GetField("activePreview", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
                form.Close();
                Assert(active.IsDisposed, "Main close left a preview orphaned.");
            }
            File.WriteAllText(Path.Combine(output, "UITestResults.txt"), "PASS: startup/title/footer, responsive compact/large main and preview windows, bounded dropdowns, busy chrome, minimize/restore retains zoom, wheel/drag/zoom slider/buttons/Fit/100%, missing-photo navigation clears stale image, view zoom leaves merge geometry unchanged, independent modeless preview lifecycle and main-close cleanup, text repair/separate-photo UI merge with exact report file/card association and source preservation, Restore Defaults preserves files and resets new options. Platform: " + Environment.OSVersion.Platform + ". Automated checks exercise window sizes; physical Windows taskbar/maximize behavior and Word require target-machine acceptance checks.");
        }
    }
}
