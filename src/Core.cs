using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Threading.Tasks;

namespace LetterMerger
{
    public sealed class CsvColumn
    {
        public int Position;
        public string Original, Label, Status;
    }

    public sealed class CsvRow
    {
        public int Record;
        public string[] Values;
    }

    public sealed class CsvData
    {
        public List<CsvColumn> Columns = new List<CsvColumn>();
        public List<CsvRow> Rows = new List<CsvRow>();
    }

    public sealed class Overlay
    {
        public double Width = 2.10, Height = 2.60, X = 0, Y = 0;
    }

    public sealed class Selection
    {
        public int Start, End, Count;
    }

    public sealed class PhotoBox
    {
        public double Width, Height, FrameWidth, FrameHeight, X, Y;
    }

    public sealed class PhotoPlan
    {
        public double SourceX, SourceY, CropWidth, CropHeight, HorizontalCrop, VerticalCrop;
        public int RenderWidth, RenderHeight;
        public long DisplayWidth, DisplayHeight, X, Y;
    }

    public sealed class PhotoData
    {
        public byte[] Bytes;
        public PhotoPlan Plan;
        public string Decoder;
        public int Frames;
    }

    public sealed class LoadedPhoto : IDisposable
    {
        public Image Image;
        public string Decoder;
        public int Frames = 1;
        public void Dispose()
        {
            if (Image != null)
                Image.Dispose();
        }
    }

    public sealed class Student
    {
        public int Record, DataRow, Card, Words;
        public string Name, Scholarship, Message, Photo, PhotoPath, PhotoStatus = "PHOTO_OK", PhotoDetails = "";
        public bool Selected, PhotoAdded, EncodingRepairEnabled;
        public string SourceMessage, OutputFile = "";
        public List<TextChange> TextChanges = new List<TextChange>();
        public List<string> TextIssues = new List<string>();
        public double Font;
        public PhotoPlan Plan;
    }

    public sealed class CardTemplate
    {
        public XmlDocument Document, Relationships, Types;
        public XmlNamespaceManager Ns;
        public List<XmlNode> Prefix;
        public XmlNode Section, MessageStyle;
        public int NameIndex, ScholarshipIndex, FrameIndex, SlotIndex;
        public string Source, Hash;
        public double Fold;
    }

    public static class Core
    {
        public const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", A = "http://schemas.openxmlformats.org/drawingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", W14 = "http://schemas.microsoft.com/office/word/2010/wordml", WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture", XML = "http://www.w3.org/XML/1998/namespace", MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static readonly string[] Modes =
        {
            "Fit entire photo",
            "Fill frame",
            "Overlay curved corners"
        };
        public static readonly string[] PhotoExtensions =
        {
            ".jpg",
            ".jpeg",
            ".png",
            ".bmp",
            ".tif",
            ".tiff",
            ".gif",
            ".jfif",
            ".jpe",
            ".webp",
            ".heic",
            ".heif"
        };
        public static string Hash(string path)
        {
            using (var h = SHA256.Create())
            using (var f = File.OpenRead(path))
                return BitConverter.ToString(h.ComputeHash(f)).Replace("-", "").ToLowerInvariant();
        }

        public static string Canonical(string value)
        {
            return Regex.Replace(System.Net.WebUtility.HtmlDecode(Regex.Replace(value ?? "", "<[^>]*>", " ")).ToLowerInvariant(), @"[^\p{L}\p{Nd}]", "");
        }

        public static string Normalize(string s)
        {
            return Regex.Replace((s ?? "").Replace("\u00ad", ""), @"\s+", "");
        }

        public static string PhotoOwner(string name)
        {
            return (name ?? "").Trim().Normalize(NormalizationForm.FormC);
        }

        public static Selection Select(int total, int start, int count, bool all)
        {
            if (total < 1)
                throw new InvalidDataException("The CSV has no nonblank data rows.");
            if (all)
                return new Selection
                {
                    Start = 1,
                    End = total,
                    Count = total
                };
            if (start < 1 || start > total || count < 1)
                throw new InvalidDataException("Choose a start row within the CSV and a positive row count.");
            int end = (int)Math.Min((long)total, (long)start + count - 1);
            return new Selection
            {
                Start = start,
                End = end,
                Count = end - start + 1
            };
        }

        public static List<string[]> ParseCsv(string text)
        {
            var rows = new List<string[]>();
            var fields = new List<string>();
            var field = new StringBuilder();
            bool quoted = false, closed = false, started = false;
            int record = 1;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (quoted)
                {
                    if (c == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"')
                        {
                            field.Append('"');
                            i++;
                        }
                        else
                        {
                            quoted = false;
                            closed = true;
                        }
                    }
                    else
                        field.Append(c);
                    continue;
                }

                if (c == '"')
                {
                    if (field.Length > 0 || started || closed)
                        throw new InvalidDataException("Unexpected quote in CSV record " + record + ".");
                    quoted = true;
                    started = true;
                    continue;
                }

                if (c == ',' || c == '\r' || c == '\n')
                {
                    fields.Add(field.ToString());
                    field.Clear();
                    closed = false;
                    started = false;
                    if (c != ',')
                    {
                        if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                            i++;
                        rows.Add(fields.ToArray());
                        fields.Clear();
                        record++;
                    }

                    continue;
                }

                if (closed)
                    throw new InvalidDataException("Unexpected text after closing quote in CSV record " + record + ".");
                field.Append(c);
                started = true;
            }

            if (quoted)
                throw new InvalidDataException("Unclosed quoted field in CSV record " + record + ".");
            if (started || closed || field.Length > 0 || fields.Count > 0)
            {
                fields.Add(field.ToString());
                rows.Add(fields.ToArray());
            }

            return rows;
        }

        public static CsvData ReadCsv(string path)
        {
            RejectLink(path);
            if (new FileInfo(path).Length > 64L * 1024 * 1024)
                throw new InvalidDataException("CSV exceeds 64 MB.");
            byte[] b = File.ReadAllBytes(path);
            if ((b.Length >= 4 && b[0] == 80 && b[1] == 75 && b[2] == 3 && b[3] == 4) || (b.Length >= 2 && b[0] == 208 && b[1] == 207))
                throw new InvalidDataException("This is an Excel workbook. In Excel, Save As CSV UTF-8 (Comma delimited). Renaming XLSX does not convert it.");
            string text;
            using (var sr = new StreamReader(new MemoryStream(b), new UTF8Encoding(false, true), true))
                text = sr.ReadToEnd();
            var records = ParseCsv(text);
            if (records.Count == 0 || records[0].Length < 4 || records[0].All(String.IsNullOrWhiteSpace))
                throw new InvalidDataException("Put at least four actual column headings in the first CSV record.");
            var data = new CsvData();
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in records[0])
            {
                string h = value.Trim();
                if (h.Length == 0)
                    continue;
                reserved.Add(h);
                counts[h] = counts.ContainsKey(h) ? counts[h] + 1 : 1;
            }

            for (int i = 0; i < records[0].Length; i++)
            {
                string original = records[0][i], label = original.Trim(), status = "ORIGINAL";
                if (label.Length == 0)
                {
                    label = "[CSV column " + (i + 1) + "] Unnamed";
                    status = "BLANK_HEADER";
                }
                else if (counts[label] > 1)
                {
                    label = "[CSV column " + (i + 1) + "] " + label;
                    status = "REPEATED_HEADER";
                }
                else if (label != original)
                    status = "TRIMMED_HEADER";
                if (status == "BLANK_HEADER" || status == "REPEATED_HEADER")
                {
                    string root = label;
                    int suffix = 1;
                    while (reserved.Contains(label) || used.Contains(label))
                        label = root + " [" + (++suffix) + "]";
                }

                used.Add(label);
                data.Columns.Add(new CsvColumn { Position = i + 1, Original = original, Label = label, Status = status });
            }

            for (int i = 1; i < records.Count; i++)
            {
                if (records[i].All(String.IsNullOrWhiteSpace))
                    continue;
                if (records[i].Length != data.Columns.Count)
                    throw new InvalidDataException("CSV record " + (i + 1) + " has " + records[i].Length + " columns; expected " + data.Columns.Count + ".");
                data.Rows.Add(new CsvRow { Record = i + 1, Values = records[i] });
            }

            return data;
        }

        public static void RejectLink(string path)
        {
            string full = Path.GetFullPath(path);
            FileSystemInfo info = File.Exists(full) ? (FileSystemInfo)new FileInfo(full) : (FileSystemInfo)new DirectoryInfo(full);
            while (info != null)
            {
                if (info.Exists && (info.Attributes & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidDataException("File/folder links are unsupported. Use extracted local files.");
                info = info is FileInfo ? ((FileInfo)info).Directory : ((DirectoryInfo)info).Parent;
            }
        }

        public static string PhotoPath(string root, string raw)
        {
            raw = (raw ?? "").Trim().Replace('\\', '/');
            if (raw.Length == 0)
                throw new InvalidDataException("EMPTY_PHOTO_PATH: the photo-path cell is blank.");
            if (Path.IsPathRooted(raw) || raw.Contains(":") || raw.Any(c => c < 32) || raw.Split('/').Contains(".."))
                throw new InvalidDataException("Use the exported local relative photo path, without URLs or parent segments.");
            string prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string full = Path.GetFullPath(Path.Combine(prefix, raw.Replace('/', Path.DirectorySeparatorChar)));
            if (!full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Photo path is outside Photos.");
            return full;
        }

        public static string ResolvePhoto(string root, string raw)
        {
            string full = PhotoPath(root, raw);
            if (!PhotoExtensions.Contains(Path.GetExtension(full).ToLowerInvariant()))
                throw new InvalidDataException("UNSUPPORTED_PHOTO_FORMAT: PDF and Word documents are not photos.");
            if (!File.Exists(full))
                throw new FileNotFoundException("PHOTO_NOT_FOUND: preserve files/documents/numbers beneath Photos.");
            RejectLink(full);
            using (var p = OpenPhoto(full))
            {
            }

            return full;
        }

        static Type WinType(string name)
        {
            return Type.GetType(name + ", Windows, ContentType=WindowsRuntime", true);
        }

        static object Call(Type type, object instance, string name, params object[] args)
        {
            var methods = type.GetMethods().Where(m => m.Name == name && m.GetParameters().Length == args.Length);
            foreach (var method in methods)
            {
                try
                {
                    return method.Invoke(instance, args);
                }
                catch (ArgumentException)
                {
                    continue;
                }
                catch (TargetInvocationException e)
                {
                    throw e.InnerException ?? e;
                }
            }

            throw new MissingMethodException(type.FullName, name);
        }

        static object AwaitWin(object op, Type result)
        {
            Type bridge = Type.GetType("System.WindowsRuntimeSystemExtensions, System.Runtime.WindowsRuntime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089", true);
            var method = bridge.GetMethods().First(m => m.Name == "AsTask" && m.IsGenericMethodDefinition && m.GetGenericArguments().Length == 1 && m.GetParameters().Length == 1 && m.GetParameters()[0].ParameterType.Name == "IAsyncOperation`1");
            var task = (Task)method.MakeGenericMethod(result).Invoke(null, new[] { op });
            if (!task.Wait(30000))
            {
                try
                {
                    Call(op.GetType(), op, "Cancel");
                }
                catch
                {
                }

                throw new TimeoutException("PHOTO_DECODE_TIMEOUT: Windows image operation exceeded 30 seconds.");
            }

            return task.GetType().GetProperty("Result").GetValue(task, null);
        }

        static object Prop(object x, string name)
        {
            return x.GetType().GetProperty(name).GetValue(x, null);
        }

        static LoadedPhoto WinPhoto(string path)
        {
            object stream = null, software = null;
            Bitmap bmp = null;
            BitmapData locked = null;
            try
            {
                Type storage = WinType("Windows.Storage.StorageFile"), decoderType = WinType("Windows.Graphics.Imaging.BitmapDecoder");
                object file = AwaitWin(Call(storage, null, "GetFileFromPathAsync", path), storage);
                stream = AwaitWin(Call(storage, file, "OpenAsync", Enum.Parse(WinType("Windows.Storage.FileAccessMode"), "Read")), WinType("Windows.Storage.Streams.IRandomAccessStream"));
                string idName = Path.GetExtension(path).Equals(".webp", StringComparison.OrdinalIgnoreCase) ? "WebpDecoderId" : "HeifDecoderId";
                Guid id = (Guid)decoderType.GetProperty(idName).GetValue(null, null);
                object decoder = AwaitWin(Call(decoderType, null, "CreateAsync", id, stream), decoderType);
                long pixels = Convert.ToInt64(Prop(decoder, "PixelWidth")) * Convert.ToInt64(Prop(decoder, "PixelHeight"));
                if (pixels < 1 || pixels > 64000000)
                    throw new InvalidDataException("PHOTO_TOO_LARGE: exceeds 64 million pixels.");
                object operation = Call(decoderType, decoder, "GetSoftwareBitmapAsync", Enum.Parse(WinType("Windows.Graphics.Imaging.BitmapPixelFormat"), "Bgra8"), Enum.Parse(WinType("Windows.Graphics.Imaging.BitmapAlphaMode"), "Straight"), Activator.CreateInstance(WinType("Windows.Graphics.Imaging.BitmapTransform")), Enum.Parse(WinType("Windows.Graphics.Imaging.ExifOrientationMode"), "RespectExifOrientation"), Enum.Parse(WinType("Windows.Graphics.Imaging.ColorManagementMode"), "ColorManageToSRgb"));
                software = AwaitWin(operation, WinType("Windows.Graphics.Imaging.SoftwareBitmap"));
                int width = Convert.ToInt32(Prop(software, "PixelWidth")), height = Convert.ToInt32(Prop(software, "PixelHeight"));
                if (width < 1 || height < 1 || (long)width * height > 64000000)
                    throw new InvalidDataException("PHOTO_TOO_LARGE_OR_INVALID.");
                byte[] bytes = new byte[checked(width * height * 4)];
                Type buffer = Type.GetType("System.Runtime.InteropServices.WindowsRuntime.WindowsRuntimeBufferExtensions, System.Runtime.WindowsRuntime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089", true);
                object winBuffer = Call(buffer, null, "AsBuffer", (object)bytes);
                Call(software.GetType(), software, "CopyToBuffer", winBuffer);
                bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                locked = bmp.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                for (int y = 0; y < height; y++)
                    Marshal.Copy(bytes, y * width * 4, IntPtr.Add(locked.Scan0, y * locked.Stride), width * 4);
                bmp.UnlockBits(locked);
                locked = null;
                var result = new LoadedPhoto
                {
                    Image = bmp,
                    Decoder = "Windows image codec",
                    Frames = Convert.ToInt32(Prop(decoder, "FrameCount"))
                };
                bmp = null;
                return result;
            }
            catch (Exception e)
            {
                throw new InvalidDataException("CODEC_OR_PHOTO_ERROR: " + Path.GetExtension(path) + " could not be decoded using the installed Windows codec. Convert it to JPG/PNG if the codec is unavailable. " + e.GetBaseException().Message, e);
            }
            finally
            {
                if (bmp != null)
                {
                    if (locked != null)
                        bmp.UnlockBits(locked);
                    bmp.Dispose();
                }

                if (software is IDisposable)
                    ((IDisposable)software).Dispose();
                if (stream is IDisposable)
                    ((IDisposable)stream).Dispose();
            }
        }

        public static LoadedPhoto OpenPhoto(string path)
        {
            RejectLink(path);
            if (new FileInfo(path).Length > 64L * 1024 * 1024)
                throw new InvalidDataException("PHOTO_TOO_LARGE: file exceeds 64 MB.");
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (!PhotoExtensions.Contains(ext))
                throw new InvalidDataException("UNSUPPORTED_PHOTO_FORMAT.");
            if (ext == ".webp" || ext == ".heic" || ext == ".heif")
                return WinPhoto(path);
            Image image = null;
            try
            {
                var header = new byte[8];
                int read;
                using (var stream = File.OpenRead(path))
                    read = stream.Read(header, 0, header.Length);
                bool raster = read >= 3 && header[0] == 255 && header[1] == 216 && header[2] == 255;
                raster |= read == 8 && header.SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
                raster |= read >= 6 && (Encoding.ASCII.GetString(header, 0, 6) == "GIF87a" || Encoding.ASCII.GetString(header, 0, 6) == "GIF89a");
                raster |= read >= 2 && header[0] == 66 && header[1] == 77;
                raster |= read >= 4 && (header.Take(4).SequenceEqual(new byte[] { 73, 73, 42, 0 }) || header.Take(4).SequenceEqual(new byte[] { 77, 77, 0, 42 }));
                if (!raster)
                    throw new InvalidDataException("UNSUPPORTED_PHOTO_CONTENT: use a raster JPG, PNG, BMP, GIF or TIFF image. Vector/metafile content is unsupported.");
                image = Image.FromFile(path);
                if (!new[] { ImageFormat.Jpeg.Guid, ImageFormat.Png.Guid, ImageFormat.Bmp.Guid, ImageFormat.Gif.Guid, ImageFormat.Tiff.Guid }.Contains(image.RawFormat.Guid))
                    throw new InvalidDataException("UNSUPPORTED_PHOTO_CONTENT: decoded content is not a supported raster image.");
                int frames = 1;
                var dims = image.FrameDimensionsList;
                var dim = dims.Contains(FrameDimension.Page.Guid) ? FrameDimension.Page : dims.Contains(FrameDimension.Time.Guid) ? FrameDimension.Time : dims.Length > 0 ? new FrameDimension(dims[0]) : null;
                if (dim != null)
                {
                    frames = image.GetFrameCount(dim);
                    image.SelectActiveFrame(dim, 0);
                }

                if (image.Width < 1 || image.Height < 1 || (long)image.Width * image.Height > 64000000)
                    throw new InvalidDataException("PHOTO_TOO_LARGE_OR_INVALID: exceeds 64 million pixels.");
                if (image.PropertyIdList.Contains(274))
                {
                    var property = image.GetPropertyItem(274);
                    if (property.Value.Length >= 2)
                    {
                        int orientation = BitConverter.ToUInt16(property.Value, 0);
                        int[] rotations =
                        {
                            0,
                            0,
                            4,
                            2,
                            6,
                            5,
                            1,
                            7,
                            3
                        };
                        if (orientation >= 1 && orientation <= 8)
                            image.RotateFlip((RotateFlipType)rotations[orientation]);
                    }
                }

                var loaded = new LoadedPhoto
                {
                    Image = image,
                    Decoder = "Windows GDI+",
                    Frames = frames
                };
                image = null;
                return loaded;
            }
            finally
            {
                if (image != null)
                    image.Dispose();
            }
        }

        public static PhotoBox Box(CardTemplate t, string mode, Overlay o)
        {
            if (!Modes.Contains(mode))
                throw new InvalidDataException("Select a valid photo sizing mode.");
            var anchor = t.Prefix[t.FrameIndex].SelectSingleNode(".//wp:anchor[.//a:blip]", t.Ns);
            var ext = (XmlElement)anchor.SelectSingleNode("wp:extent", t.Ns);
            double fw = Double.Parse(ext.GetAttribute("cx"), CultureInfo.InvariantCulture), fh = Double.Parse(ext.GetAttribute("cy"), CultureInfo.InvariantCulture);
            var box = new PhotoBox
            {
                FrameWidth = fw,
                FrameHeight = fh,
                Width = Math.Min(1.9 * 914400, fw - .49 * 914400),
                Height = Math.Min(2.4 * 914400, fh - .57 * 914400)
            };
            if (mode == Modes[2])
            {
                o = o ?? new Overlay();
                box.Width = o.Width * 914400;
                box.Height = o.Height * 914400;
                box.X = o.X * 914400;
                box.Y = o.Y * 914400;
            }

            if (new[]
            {
                box.Width,
                box.Height,
                box.X,
                box.Y
            }.Any(v => Double.IsNaN(v) || Double.IsInfinity(v)) || box.Width <= 0 || box.Height <= 0 || box.Width + 2 * Math.Abs(box.X) > fw || box.Height + 2 * Math.Abs(box.Y) > fh)
                throw new InvalidDataException("Photo dimensions and position must stay inside the original full frame. Reduce the size or center it.");
            return box;
        }

        public static PhotoPlan Plan(double width, double height, PhotoBox box, string mode)
        {
            if (!Modes.Contains(mode) || width <= 0 || height <= 0)
                throw new InvalidDataException("Invalid photo dimensions/mode.");
            double cropWidth = width, cropHeight = height, displayWidth, displayHeight;
            if (mode != Modes[0])
            {
                if (width / height > box.Width / box.Height)
                    cropWidth = height * box.Width / box.Height;
                else
                    cropHeight = width * box.Height / box.Width;
                displayWidth = box.Width;
                displayHeight = box.Height;
            }
            else
            {
                double fit = Math.Min(box.Width / width, box.Height / height);
                displayWidth = width * fit;
                displayHeight = height * fit;
            }

            double scale = Math.Min(1, Math.Min(displayWidth / 914400 * 300 / cropWidth, displayHeight / 914400 * 300 / cropHeight));
            return new PhotoPlan
            {
                SourceX = (width - cropWidth) / 2,
                SourceY = (height - cropHeight) / 2,
                CropWidth = cropWidth,
                CropHeight = cropHeight,
                RenderWidth = Math.Max(1, (int)Math.Round(cropWidth * scale)),
                RenderHeight = Math.Max(1, (int)Math.Round(cropHeight * scale)),
                DisplayWidth = (long)Math.Round(displayWidth),
                DisplayHeight = (long)Math.Round(displayHeight),
                X = (long)Math.Round(box.X),
                Y = (long)Math.Round(box.Y),
                HorizontalCrop = Math.Round(100 * (1 - cropWidth / width), 2),
                VerticalCrop = Math.Round(100 * (1 - cropHeight / height), 2)
            };
        }

        public static PhotoData ReadPhoto(string path, PhotoBox box, string mode)
        {
            using (var loaded = OpenPhoto(path))
            {
                PhotoPlan plan = Plan(loaded.Image.Width, loaded.Image.Height, box, mode);
                using (var bitmap = new Bitmap(plan.RenderWidth, plan.RenderHeight))
                using (var g = Graphics.FromImage(bitmap))
                using (var memory = new MemoryStream())
                {
                    bitmap.SetResolution(300, 300);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.Clear(Color.White);
                    g.DrawImage(loaded.Image, new RectangleF(0, 0, plan.RenderWidth, plan.RenderHeight), new RectangleF((float)plan.SourceX, (float)plan.SourceY, (float)plan.CropWidth, (float)plan.CropHeight), GraphicsUnit.Pixel);
                    bitmap.Save(memory, ImageFormat.Png);
                    return new PhotoData
                    {
                        Bytes = memory.ToArray(),
                        Plan = plan,
                        Decoder = loaded.Decoder,
                        Frames = loaded.Frames
                    };
                }
            }
        }

        public static XmlNamespaceManager Names(XmlDocument d)
        {
            var ns = new XmlNamespaceManager(d.NameTable);
            foreach (var p in new[]
            {
                new[]
                {
                    "w",
                    W
                },
                new[]
                {
                    "wp",
                    WP
                },
                new[]
                {
                    "a",
                    A
                },
                new[]
                {
                    "r",
                    R
                },
                new[]
                {
                    "w14",
                    W14
                },
                new[]
                {
                    "wp14",
                    WP14
                },
                new[]
                {
                    "pic",
                    PIC
                }
            }

            )
                ns.AddNamespace(p[0], p[1]);
            return ns;
        }

        public static void CheckArchive(ZipArchive z, bool mergedOutput = false)
        {
            if (z.Entries.Count > 5000)
                throw new InvalidDataException("DOCX has too many parts.");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            foreach (var e in z.Entries)
            {
                string name = e.FullName;
                if (name.StartsWith("/") || name.Contains("\\") || name.Split('/').Any(p => p == ".." || p == ".") || !names.Add(name) || name.Contains(":") || name.Any(c => c < 32))
                    throw new InvalidDataException("Unsafe or repeated DOCX part name.");
                total += e.Length;
                if (e.Length > 64L * 1024 * 1024 || total > (mergedOutput ? 512L : 128L) * 1024 * 1024)
                    throw new InvalidDataException("DOCX exceeds the supported uncompressed size.");
                if (Regex.IsMatch(name, @"(^|/)(vbaProject\.bin|activeX/|embeddings/)|\.(exe|dll|com|scr|ps1|bat|cmd|vbs|js|hta)$", RegexOptions.IgnoreCase))
                    throw new InvalidDataException("Executable/embedded-object content is unsupported. Use the supplied A7 DOCX template.");
            }
        }

        public static XmlDocument ReadXml(ZipArchive z, string name)
        {
            var e = z.GetEntry(name);
            if (e == null)
                throw new InvalidDataException("Missing DOCX part: " + name);
            var doc = new XmlDocument
            {
                PreserveWhitespace = true,
                XmlResolver = null
            };
            using (var s = e.Open())
            using (var reader = XmlReader.Create(s, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64L * 1024 * 1024 }))
                doc.Load(reader);
            return doc;
        }

        public static byte[] XmlBytes(XmlDocument doc)
        {
            using (var m = new MemoryStream())
            {
                using (var writer = XmlWriter.Create(m, new XmlWriterSettings { Encoding = new UTF8Encoding(false, true), Indent = false, CheckCharacters = true, CloseOutput = false }))
                {
                    doc.Save(writer);
                    writer.Flush();
                }

                return m.ToArray();
            }
        }

        public static string Text(XmlNode n, XmlNamespaceManager ns)
        {
            return String.Concat(n.SelectNodes(".//w:t", ns).Cast<XmlNode>().Select(x => x.InnerText));
        }

        static string Instructions(XmlNode n, XmlNamespaceManager ns)
        {
            return String.Concat(n.SelectNodes(".//w:instrText", ns).Cast<XmlNode>().Select(x => x.InnerText));
        }

        static void CheckTemplateContent(ZipArchive archive, XmlDocument types)
        {
            foreach (XmlElement type in types.DocumentElement.ChildNodes.OfType<XmlElement>())
                if (Regex.IsMatch(type.GetAttribute("ContentType"), @"macroEnabled|vbaProject|activeX|oleObject", RegexOptions.IgnoreCase))
                    throw new InvalidDataException("Macro/active-object content is unsupported in templates.");
            foreach (var part in archive.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                var document = ReadXml(archive, part.FullName);
                var ns = Names(document);
                var instructions = document.SelectNodes("//w:instrText|//w:fldSimple/@w:instr", ns).Cast<XmlNode>().Select(n => n.InnerText).ToList();
                string combined = String.Concat(instructions);
                if (Regex.IsMatch(combined, @"\b(DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|LINK|DATABASE|MACROBUTTON)\b", RegexOptions.IgnoreCase))
                    throw new InvalidDataException("Active/external Word fields are unsupported in template part " + part.FullName + ".");
                if (document.SelectNodes("//*[local-name()='altChunk']").Count > 0)
                    throw new InvalidDataException("Imported document chunks are unsupported.");
            }
        }

        public static CardTemplate ReadTemplate(string path)
        {
            RejectLink(path);
            if (Path.GetExtension(path).ToLowerInvariant() != ".docx")
                throw new InvalidDataException("Use a DOCX template.");
            if (new FileInfo(path).Length > 64L * 1024 * 1024)
                throw new InvalidDataException("Template exceeds 64 MB.");
            var t = new CardTemplate
            {
                Source = path,
                Hash = Hash(path)
            };
            using (var z = ZipFile.OpenRead(path))
            {
                CheckArchive(z);
                t.Document = ReadXml(z, "word/document.xml");
                t.Relationships = ReadXml(z, "word/_rels/document.xml.rels");
                t.Types = ReadXml(z, "[Content_Types].xml");
                CheckTemplateContent(z, t.Types);
                Security.CheckTemplate(z);
                foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".rels")))
                {
                    var rel = ReadXml(z, e.FullName);
                    foreach (XmlElement node in rel.DocumentElement.ChildNodes.OfType<XmlElement>())
                        if (node.GetAttribute("TargetMode") == "External")
                            throw new InvalidDataException("External links/resources in templates are unsupported. Use the supplied A7 template.");
                }

                if (t.Document.SelectNodes("//*[local-name()='altChunk']").Count > 0)
                    throw new InvalidDataException("Imported document chunks are unsupported.");
            }

            t.Ns = Names(t.Document);
            foreach (XmlNode instruction in t.Document.SelectNodes("//w:instrText|//w:fldSimple/@w:instr", t.Ns))
                if (Regex.IsMatch(instruction.InnerText, @"\b(DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|LINK|DATABASE|MACROBUTTON)\b", RegexOptions.IgnoreCase))
                    throw new InvalidDataException("Active/external Word fields are unsupported.");
            var paragraphs = t.Document.SelectNodes("/w:document/w:body/w:p", t.Ns).Cast<XmlNode>().ToList();
            int label = -1, fold = -1, slot = -1, message = -1;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = Text(paragraphs[i], t.Ns);
                if (text.Trim() == "IMAGE")
                    slot = i;
                if (text.StartsWith("Recipient of the", StringComparison.Ordinal))
                    label = i;
                if (label >= 0 && i > label && paragraphs[i].SelectSingleNode(".//wp:anchor", t.Ns) != null)
                    fold = i;
                if (Regex.IsMatch(Instructions(paragraphs[i], t.Ns), @"MERGEFIELD\s+Please_draft_a_Thank_You_letter"))
                    message = i;
            }

            if (label < 1 || label + 1 >= paragraphs.Count || fold < 0 || slot < 0 || message <= fold || slot >= label - 1 || t.Document.SelectNodes("//wp:anchor[.//a:blip]", t.Ns).Count != 1 || !Regex.IsMatch(Instructions(paragraphs[label - 1], t.Ns), @"MERGEFIELD\s+Name\b") || !Regex.IsMatch(Instructions(paragraphs[label + 1], t.Ns), @"MERGEFIELD\s+Portfolio_Name\b"))
                throw new InvalidDataException("Use the supplied original A7_Card_Template.docx with Name, Portfolio_Name, IMAGE, fold line and thank-you fields.");
            var section = t.Document.SelectSingleNode("/w:document/w:body/w:sectPr", t.Ns);
            if (section == null || section.SelectSingleNode("w:pgSz", t.Ns) == null || section.SelectSingleNode("w:pgMar", t.Ns) == null)
                throw new InvalidDataException("Template page size/margins are missing.");
            t.Section = section.CloneNode(true);
            t.Prefix = paragraphs.Take(fold + 1).Select(n => n.CloneNode(true)).ToList();
            t.FrameIndex = t.Prefix.FindIndex(p => p.SelectSingleNode(".//wp:anchor[.//a:blip]", t.Ns) != null);
            t.NameIndex = label - 1;
            t.ScholarshipIndex = label + 1;
            t.SlotIndex = slot;
            t.MessageStyle = paragraphs[message].CloneNode(true);
            var position = (XmlElement)paragraphs[fold].SelectSingleNode(".//wp:anchor/wp:positionV", t.Ns);
            if (position == null || position.GetAttribute("relativeFrom") != "page" || position.SelectSingleNode("wp:posOffset", t.Ns) == null)
                throw new InvalidDataException("The fold line must retain its fixed page position.");
            t.Fold = Double.Parse(position.SelectSingleNode("wp:posOffset", t.Ns).InnerText, CultureInfo.InvariantCulture) / 12700;
            return t;
        }

        public static void Attribute(XmlElement e, string prefix, string name, string uri, string value)
        {
            XmlAttribute a = e.GetAttributeNode(name, uri);
            if (a == null)
            {
                a = e.OwnerDocument.CreateAttribute(prefix, name, uri);
                e.Attributes.Append(a);
            }

            a.Prefix = prefix;
            a.Value = value;
        }

        static XmlElement Element(XmlDocument d, string name)
        {
            return d.CreateElement("w", name, W);
        }

        static readonly Dictionary<string, string> Orders = new Dictionary<string, string>
        {
            {
                "pPr",
                "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange"
            },
            {
                "rPr",
                "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange"
            },
            {
                "sectPr",
                "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange"
            }
        };
        public static XmlElement Property(XmlNode parent, string name, params string[] attrs)
        {
            var node = parent.ChildNodes.OfType<XmlElement>().FirstOrDefault(n => n.LocalName == name && n.NamespaceURI == W);
            if (node == null)
            {
                node = Element(parent.OwnerDocument, name);
                XmlNode before = null;
                if (Orders.ContainsKey(parent.LocalName))
                {
                    var order = Orders[parent.LocalName].Split(' ');
                    int rank = Array.IndexOf(order, name);
                    if (rank >= 0)
                        before = parent.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => Array.IndexOf(order, n.LocalName) > rank);
                }

                if (before != null)
                    parent.InsertBefore(node, before);
                else
                    parent.AppendChild(node);
            }

            for (int i = 0; i < attrs.Length; i += 2)
                Attribute(node, "w", attrs[i], W, attrs[i + 1]);
            return node;
        }

        public static void Fill(XmlNode p, string value, XmlNamespaceManager ns)
        {
            XmlNode format = p.SelectSingleNode("w:r[w:t]/w:rPr", ns);
            if (format != null)
                format = format.CloneNode(true);
            foreach (XmlNode child in p.ChildNodes.Cast<XmlNode>().ToArray())
                if (child.LocalName != "pPr")
                    p.RemoveChild(child);
            var run = Element(p.OwnerDocument, "r");
            if (format != null)
                run.AppendChild(format);
            var text = Element(p.OwnerDocument, "t");
            Attribute(text, "xml", "space", XML, "preserve");
            text.InnerText = value;
            run.AppendChild(text);
            p.AppendChild(run);
        }

        public static List<string> BodyLines(string raw, string name)
        {
            var lines = raw.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').Where(s => !String.IsNullOrWhiteSpace(s)).ToList();
            var output = new List<string>();
            if (lines.Count == 0)
                return output;
            if (!Regex.IsMatch(lines[0].Trim(), @"^(Dear\b|Hello\b|To\b|Greetings\b)", RegexOptions.IgnoreCase))
                output.Add("Dear Donor,");
            output.AddRange(lines);
            string pattern = @"^(sincerely|respectfully|very respectfully|best regards|kind regards|warm regards|with .*gratitude|with .*appreciation|thank you)[,!.: ]*$";
            bool closing = lines.Skip(Math.Max(0, lines.Count - 8)).Any(l => Regex.IsMatch(l.Trim(' ', '_', '*'), pattern, RegexOptions.IgnoreCase));
            if (!closing)
            {
                output.Add("Sincerely,");
                output.Add(name);
            }
            else if (Regex.IsMatch(lines.Last().Trim(' ', '_', '*'), pattern, RegexOptions.IgnoreCase))
                output.Add(name);
            return output;
        }

        static XmlNode MessageParagraph(CardTemplate t, string text, double size, bool first)
        {
            var p = Element(t.Document, "p");
            var old = t.MessageStyle.SelectSingleNode("w:pPr", t.Ns);
            if (old != null)
                p.AppendChild(old.CloneNode(true));
            var pr = Property(p, "pPr");
            foreach (XmlNode n in pr.SelectNodes("w:sectPr|w:ind|w:tabs|w:jc", t.Ns).Cast<XmlNode>().ToArray())
                pr.RemoveChild(n);
            Property(pr, "spacing", "before", first ? "360" : "0", "after", "50", "line", ((int)Math.Round(Math.Max(11.8, size * 1.12) * 20)).ToString(CultureInfo.InvariantCulture), "lineRule", "exact");
            foreach (string n in new[]
            {
                "keepNext",
                "keepLines",
                "widowControl",
                "snapToGrid"
            }

            )
                Property(pr, n, "val", "0");
            string fontSize = ((int)(size * 2)).ToString(CultureInfo.InvariantCulture);
            Property(Property(pr, "rPr"), "sz", "val", fontSize);
            var run = Element(t.Document, "r");
            var rpr = Property(run, "rPr");
            Property(rpr, "rFonts", "ascii", "Times New Roman", "hAnsi", "Times New Roman", "eastAsia", "Times New Roman", "cs", "Times New Roman");
            Property(rpr, "sz", "val", fontSize);
            Property(rpr, "szCs", "val", fontSize);
            var txt = Element(t.Document, "t");
            Attribute(txt, "xml", "space", XML, "preserve");
            txt.InnerText = text;
            run.AppendChild(txt);
            p.AppendChild(run);
            return p;
        }

        static void Bookmark(XmlNode first, XmlNode last, string name, int id)
        {
            var s = Element(first.OwnerDocument, "bookmarkStart");
            Attribute(s, "w", "id", W, id.ToString());
            Attribute(s, "w", "name", W, name);
            var pr = first.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.LocalName == "pPr");
            if (pr != null)
                first.InsertAfter(s, pr);
            else
                first.PrependChild(s);
            var e = Element(last.OwnerDocument, "bookmarkEnd");
            Attribute(e, "w", "id", W, id.ToString());
            last.AppendChild(e);
        }

        static string FreshId()
        {
            return Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();
        }

        static void PlacePhoto(CardTemplate t, List<XmlNode> prefix, Student s, PhotoData photo, string rid, int id)
        {
            var anchor = (XmlElement)prefix[t.FrameIndex].SelectSingleNode(".//wp:anchor[.//a:blip]", t.Ns).CloneNode(true);
            var extent = (XmlElement)anchor.SelectSingleNode("wp:extent", t.Ns);
            long fw = Int64.Parse(extent.GetAttribute("cx")), fh = Int64.Parse(extent.GetAttribute("cy")), cx = photo.Plan.DisplayWidth, cy = photo.Plan.DisplayHeight;
            long x = photo.Plan.X, y = photo.Plan.Y;
            if (cx <= 0 || cy <= 0 || cx + 2 * Math.Abs(x) > fw || cy + 2 * Math.Abs(y) > fh)
                throw new InvalidDataException("Photo exceeds frame.");
            extent.SetAttribute("cx", cx.ToString());
            extent.SetAttribute("cy", cy.ToString());
            var vertical = anchor.SelectSingleNode("wp:positionV/wp:posOffset", t.Ns);
            vertical.InnerText = (Int64.Parse(vertical.InnerText) + (long)Math.Round((fh - cy) / 2.0) + y).ToString();
            if (x != 0)
            {
                var horizontal = (XmlElement)anchor.SelectSingleNode("wp:positionH", t.Ns);
                var align = horizontal.SelectSingleNode("wp:align", t.Ns);
                if (align == null || align.InnerText != "center" || horizontal.GetAttribute("relativeFrom") != "margin")
                    throw new InvalidDataException("Manual position requires the supplied centered A7 frame.");
                var page = (XmlElement)t.Section.SelectSingleNode("w:pgSz", t.Ns);
                var margins = (XmlElement)t.Section.SelectSingleNode("w:pgMar", t.Ns);
                double usable = (Double.Parse(page.GetAttribute("w", W)) - Double.Parse(margins.GetAttribute("left", W)) - Double.Parse(margins.GetAttribute("right", W))) * 635;
                horizontal.RemoveChild(align);
                var offset = t.Document.CreateElement("wp", "posOffset", WP);
                offset.InnerText = ((long)Math.Round((usable - cx) / 2) + x).ToString();
                horizontal.AppendChild(offset);
            }

            anchor.SetAttribute("behindDoc", "0");
            anchor.SetAttribute("relativeHeight", "251658241");
            Attribute(anchor, "wp14", "anchorId", WP14, FreshId());
            Attribute(anchor, "wp14", "editId", WP14, FreshId());
            var dp = (XmlElement)anchor.SelectSingleNode("wp:docPr", t.Ns);
            dp.SetAttribute("id", id.ToString());
            dp.SetAttribute("name", "StudentPhoto_" + s.Card);
            dp.SetAttribute("descr", "Student photograph for " + s.Name);
            Attribute((XmlElement)anchor.SelectSingleNode(".//a:blip", t.Ns), "r", "embed", R, rid);
            var cpr = (XmlElement)anchor.SelectSingleNode(".//pic:cNvPr", t.Ns);
            if (cpr != null)
            {
                cpr.SetAttribute("id", id.ToString());
                cpr.SetAttribute("name", "StudentPhoto_" + s.Card + ".png");
            }

            var inner = (XmlElement)anchor.SelectSingleNode(".//pic:spPr/a:xfrm/a:ext", t.Ns);
            if (inner != null)
            {
                inner.SetAttribute("cx", cx.ToString());
                inner.SetAttribute("cy", cy.ToString());
            }

            foreach (XmlNode crop in anchor.SelectNodes(".//a:srcRect", t.Ns).Cast<XmlNode>().ToArray())
                crop.ParentNode.RemoveChild(crop);
            var run = Element(t.Document, "r");
            var drawing = Element(t.Document, "drawing");
            drawing.AppendChild(anchor);
            run.AppendChild(drawing);
            prefix[t.FrameIndex].AppendChild(run);
            foreach (XmlNode text in prefix[t.SlotIndex].SelectNodes(".//w:t", t.Ns))
                text.InnerText = "";
        }

        public static List<Student> Build(CardTemplate t, List<Student> students, string output, string mode, Overlay overlay, string validationFile = null)
        {
            PhotoBox box = Box(t, mode, overlay);
            var body = t.Document.SelectSingleNode("/w:document/w:body", t.Ns);
            body.RemoveAll();
            var changes = Security.OutputMetadata();
            var cache = new Dictionary<string, Tuple<PhotoData, string>>(StringComparer.OrdinalIgnoreCase);
            int drawId = 1;
            for (int i = 0; i < students.Count; i++)
            {
                var s = students[i];
                s.Card = i + 1;
                s.Words = Regex.Matches(s.Message, @"\S+").Count;
                s.Font = s.Words > 230 ? 10.5 : s.Words > 150 ? 11 : s.Words > 80 ? 12 : 14;
                var prefix = t.Prefix.Select(n => n.CloneNode(true)).ToList();
                Fill(prefix[t.NameIndex], s.Name, t.Ns);
                Fill(prefix[t.ScholarshipIndex], s.Scholarship, t.Ns);
                foreach (XmlNode p in prefix)
                {
                    foreach (XmlElement anchor in p.SelectNodes(".//wp:anchor", t.Ns))
                    {
                        ((XmlElement)anchor.SelectSingleNode("wp:docPr", t.Ns)).SetAttribute("id", (drawId++).ToString());
                        Attribute(anchor, "wp14", "anchorId", WP14, FreshId());
                        Attribute(anchor, "wp14", "editId", WP14, FreshId());
                    }

                    foreach (XmlElement legacy in p.SelectNodes(".//*[namespace-uri()='urn:schemas-microsoft-com:vml' and (@id or @*[local-name()='spid'])]"))
                    {
                        if (legacy.HasAttribute("id"))
                            legacy.SetAttribute("id", "A7Shape_" + Guid.NewGuid().ToString("N"));
                        foreach (XmlAttribute attr in legacy.Attributes)
                            if (attr.LocalName == "spid")
                                attr.Value = "_x0000_s" + (2000 + i * 20 + drawId);
                    }

                    body.AppendChild(p);
                }

                Bookmark(prefix[t.NameIndex], prefix[t.NameIndex], "LM_Name_" + s.Card.ToString("D4"), 10000 + i * 3);
                Bookmark(prefix[t.ScholarshipIndex], prefix[t.ScholarshipIndex], "LM_Scholar_" + s.Card.ToString("D4"), 10001 + i * 3);
                var paragraphs = BodyLines(s.Message, s.Name).Select((line, j) => MessageParagraph(t, line, s.Font, j == 0)).ToList();
                if (paragraphs.Count == 0)
                    throw new InvalidDataException("Blank student response.");
                foreach (var p in paragraphs)
                    body.AppendChild(p);
                Bookmark(paragraphs[0], paragraphs.Last(), "LM_Message_" + s.Card.ToString("D4"), 10002 + i * 3);
                if (i < students.Count - 1)
                {
                    var sect = t.Section.CloneNode(true);
                    Property(sect, "type", "val", "nextPage");
                    Property(paragraphs.Last(), "pPr").AppendChild(sect);
                }

                s.PhotoAdded = false;
                s.Plan = null;
                if (!String.IsNullOrEmpty(s.Photo))
                {
                    try
                    {
                        if (!cache.ContainsKey(s.Photo))
                        {
                            var photo = ReadPhoto(s.Photo, box, mode);
                            string rid = "LMPhoto" + Guid.NewGuid().ToString("N"), part = "media/lm_photo_" + Guid.NewGuid().ToString("N") + ".png";
                            var relation = t.Relationships.CreateElement("Relationship", t.Relationships.DocumentElement.NamespaceURI);
                            relation.SetAttribute("Id", rid);
                            relation.SetAttribute("Type", R + "/image");
                            relation.SetAttribute("Target", part);
                            t.Relationships.DocumentElement.AppendChild(relation);
                            changes["word/" + part] = photo.Bytes;
                            cache[s.Photo] = Tuple.Create(photo, rid);
                        }

                        var item = cache[s.Photo];
                        PlacePhoto(t, prefix, s, item.Item1, item.Item2, drawId++);
                        s.PhotoAdded = true;
                        s.Plan = item.Item1.Plan;
                        s.PhotoDetails += " Decoder: " + item.Item1.Decoder + ". Mode: " + mode + ". Display " + (s.Plan.DisplayWidth / 914400.0).ToString("0.###", CultureInfo.InvariantCulture) + " x " + (s.Plan.DisplayHeight / 914400.0).ToString("0.###", CultureInfo.InvariantCulture) + " inches; offset " + (s.Plan.X / 914400.0).ToString("0.###", CultureInfo.InvariantCulture) + ", " + (s.Plan.Y / 914400.0).ToString("0.###", CultureInfo.InvariantCulture) + " inches. Crop width " + s.Plan.HorizontalCrop + "%; height " + s.Plan.VerticalCrop + "%.";
                        if (item.Item1.Frames > 1)
                            s.PhotoDetails += " FIRST_FRAME_ONLY: frame/page 1 of " + item.Item1.Frames + ".";
                    }
                    catch (Exception e)
                    {
                        s.PhotoStatus = "PHOTO_INSERT_ERROR";
                        s.PhotoDetails = e.GetBaseException().Message;
                    }
                }
            }

            var final = t.Section.CloneNode(true);
            Property(final, "type", "val", "nextPage");
            body.AppendChild(final);
            if (t.Types.SelectSingleNode("/*[local-name()='Types']/*[local-name()='Default' and @Extension='png']") == null)
            {
                var png = t.Types.CreateElement("Default", t.Types.DocumentElement.NamespaceURI);
                png.SetAttribute("Extension", "png");
                png.SetAttribute("ContentType", "image/png");
                t.Types.DocumentElement.AppendChild(png);
            }

            int paraId = 1;
            foreach (XmlElement p in t.Document.SelectNodes("//w:body//w:p", t.Ns))
            {
                Attribute(p, "w14", "paraId", W14, paraId.ToString("X8"));
                Attribute(p, "w14", "textId", W14, (paraId++).ToString("X8"));
            }

            var prefixes = new Dictionary<string, string>
            {
                {
                    W,
                    "w"
                },
                {
                    R,
                    "r"
                },
                {
                    WP14,
                    "wp14"
                },
                {
                    W14,
                    "w14"
                },
                {
                    XML,
                    "xml"
                }
            };
            foreach (XmlElement el in t.Document.SelectNodes("//*"))
            {
                foreach (XmlAttribute attr in el.Attributes)
                    if (prefixes.ContainsKey(attr.NamespaceURI))
                        attr.Prefix = prefixes[attr.NamespaceURI];
                if (el.NamespaceURI == "urn:schemas-microsoft-com:vml" && el.HasAttribute("anchorId", W14))
                    Attribute(el, "w14", "anchorId", W14, FreshId());
            }

            changes["word/document.xml"] = XmlBytes(t.Document);
            changes["word/_rels/document.xml.rels"] = XmlBytes(t.Relationships);
            changes["[Content_Types].xml"] = XmlBytes(t.Types);
            Core.RejectLink(output);
            string staging = output + ".partial";
            Core.RejectLink(staging);
            if (File.Exists(staging))
                throw new IOException("A partial output already exists. It was preserved for review.");
            try
            {
                if (Hash(t.Source) != t.Hash)
                    throw new InvalidDataException("Template changed during merge.");
                WritePackage(t.Source, staging, changes);
                if (Hash(t.Source) != t.Hash)
                    throw new InvalidDataException("Template changed during output writing.");
                string validation = Validate(staging, students);
                File.Move(staging, output);
                try
                {
                    File.WriteAllText(validationFile ?? Path.Combine(Path.GetDirectoryName(output), "DocumentValidation.txt"), validation, new UTF8Encoding(true));
                }
                catch (Exception e)
                {
                    App.LogException(e);
                }
            }
            catch
            {
                if (File.Exists(staging))
                    File.Delete(staging);
                throw;
            }

            return students;
        }

        public static void WritePackage(string source, string destination, Dictionary<string, byte[]> changes)
        {
            Core.RejectLink(source);
            Core.RejectLink(destination);
            using (var input = ZipFile.OpenRead(source))
            using (var file = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var output = new ZipArchive(file, ZipArchiveMode.Create))
            {
                foreach (var entry in input.Entries)
                {
                    var e = output.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    using (var stream = e.Open())
                    {
                        if (changes.ContainsKey(entry.FullName))
                            stream.Write(changes[entry.FullName], 0, changes[entry.FullName].Length);
                        else
                            using (var from = entry.Open())
                                from.CopyTo(stream);
                    }
                }

                foreach (var pair in changes)
                    if (input.GetEntry(pair.Key) == null)
                    {
                        var entry = output.CreateEntry(pair.Key, CompressionLevel.Optimal);
                        using (var stream = entry.Open())
                            stream.Write(pair.Value, 0, pair.Value.Length);
                    }
            }
        }

        public static string Validate(string path, List<Student> cards)
        {
            using (var z = ZipFile.OpenRead(path))
            {
                CheckArchive(z, true);
                var docs = new Dictionary<string, XmlDocument>();
                foreach (var e in z.Entries)
                    if (e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels"))
                        docs[e.FullName] = ReadXml(z, e.FullName);
                foreach (string required in new[]
                {
                    "[Content_Types].xml",
                    "_rels/.rels",
                    "word/document.xml",
                    "word/_rels/document.xml.rels"
                }

                )
                    if (!docs.ContainsKey(required))
                        throw new InvalidDataException("Missing package part " + required);
                var parts = new HashSet<string>(z.Entries.Select(e => e.FullName));
                foreach (var pair in docs.Where(p => p.Key.EndsWith(".rels")))
                {
                    string source = "";
                    if (pair.Key != "_rels/.rels")
                    {
                        var match = Regex.Match(pair.Key, @"^(.*)/_rels/([^/]+)\.rels$");
                        if (!match.Success)
                            throw new InvalidDataException("Invalid relationship part.");
                        source = match.Groups[1].Value + "/" + match.Groups[2].Value;
                    }

                    var ids = new HashSet<string>();
                    foreach (XmlElement rel in pair.Value.DocumentElement.ChildNodes.OfType<XmlElement>())
                    {
                        if (String.IsNullOrEmpty(rel.GetAttribute("Id")) || !ids.Add(rel.GetAttribute("Id")))
                            throw new InvalidDataException("Repeated/missing relationship ID.");
                        if (rel.GetAttribute("TargetMode") == "External")
                            throw new InvalidDataException("External relationship is unsupported.");
                        var uri = new Uri(new Uri("http://a7-package.invalid/" + source), rel.GetAttribute("Target"));
                        string target = Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/');
                        if (uri.Host != "a7-package.invalid" || !parts.Contains(target))
                            throw new InvalidDataException("Missing internal relationship target: " + target);
                    }
                }

                var types = docs["[Content_Types].xml"];
                CheckTemplateContent(z, types);
                Security.CheckRelationships(z);
                var defaults = new HashSet<string>();
                var overrides = new Dictionary<string, string>();
                foreach (XmlElement node in types.DocumentElement.ChildNodes.OfType<XmlElement>())
                {
                    if (node.LocalName == "Default")
                        defaults.Add(node.GetAttribute("Extension"));
                    if (node.LocalName == "Override")
                        overrides[Uri.UnescapeDataString(node.GetAttribute("PartName")).TrimStart('/')] = node.GetAttribute("ContentType");
                }

                if (!overrides.ContainsKey("word/document.xml") || overrides["word/document.xml"] != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
                    throw new InvalidDataException("Main part is not a Word DOCX document.");
                foreach (var name in parts)
                    if (name != "[Content_Types].xml" && !name.EndsWith("/") && !overrides.ContainsKey(name) && !defaults.Contains(Path.GetExtension(name).TrimStart('.')))
                        throw new InvalidDataException("Missing content type.");
                var doc = docs["word/document.xml"];
                var ns = Names(doc);
                var body = doc.SelectSingleNode("/w:document/w:body", ns);
                if (body == null || body.LastChild.LocalName != "sectPr" || body.SelectNodes("w:sectPr", ns).Count != 1)
                    throw new InvalidDataException("Final section is invalid.");
                var paragraphIds = new HashSet<string>();
                var drawingIds = new HashSet<string>();
                foreach (XmlElement node in doc.SelectNodes("//*"))
                {
                    foreach (XmlAttribute attr in node.Attributes)
                    {
                        if (attr.NamespaceURI == XML && attr.Prefix != "xml")
                            throw new InvalidDataException("Invalid XML namespace prefix.");
                        if (attr.NamespaceURI == MC && (attr.LocalName == "Ignorable" || attr.LocalName == "Requires"))
                            foreach (string p in attr.Value.Split(' '))
                                if (p.Length > 0 && String.IsNullOrEmpty(node.GetNamespaceOfPrefix(p)))
                                    throw new InvalidDataException("Undeclared compatibility prefix.");
                    }

                    if (node.LocalName == "Choice" && node.NamespaceURI == MC)
                        foreach (string p in node.GetAttribute("Requires").Split(' '))
                            if (p.Length > 0 && String.IsNullOrEmpty(node.GetNamespaceOfPrefix(p)))
                                throw new InvalidDataException("Undeclared choice prefix.");
                    if (node.LocalName == "p" && node.HasAttribute("paraId", W14))
                    {
                        string id = node.GetAttribute("paraId", W14);
                        long value;
                        if (!Regex.IsMatch(id, "^[0-9A-Fa-f]{8}$") || !Int64.TryParse(id, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value) || value <= 0 || value >= 2147483648 || !paragraphIds.Add(id))
                            throw new InvalidDataException("Invalid/repeated paragraph ID.");
                    }

                    if (node.LocalName == "docPr" && node.NamespaceURI == WP && !drawingIds.Add(node.GetAttribute("id")))
                        throw new InvalidDataException("Repeated drawing ID.");
                }

                var relations = docs["word/_rels/document.xml.rels"].DocumentElement.ChildNodes.OfType<XmlElement>().ToDictionary(e => e.GetAttribute("Id"));
                foreach (XmlAttribute attr in doc.SelectNodes("//@r:embed|//@r:id|//@r:link", ns))
                {
                    if (!relations.ContainsKey(attr.Value))
                        throw new InvalidDataException("Unresolved document relationship reference.");
                    if (attr.OwnerElement.LocalName == "blip" && relations[attr.Value].GetAttribute("Type") != R + "/image")
                        throw new InvalidDataException("Photo references a non-image part.");
                }

                var active = new Dictionary<string, Tuple<string, StringBuilder>>();
                var texts = new Dictionary<string, string>();
                var allIds = new HashSet<string>();
                foreach (XmlElement node in doc.SelectNodes("//w:bookmarkStart|//w:bookmarkEnd|//w:t", ns))
                {
                    if (node.LocalName == "bookmarkStart")
                    {
                        string id = node.GetAttribute("id", W), name = node.GetAttribute("name", W);
                        if (id.Length == 0 || name.Length == 0 || !allIds.Add(id) || texts.ContainsKey(name) || active.Values.Any(p => p.Item1 == name))
                            throw new InvalidDataException("Repeated/missing bookmark identity.");
                        active[id] = Tuple.Create(name, new StringBuilder());
                    }
                    else if (node.LocalName == "bookmarkEnd")
                    {
                        string id = node.GetAttribute("id", W);
                        if (!active.ContainsKey(id))
                            throw new InvalidDataException("Unmatched bookmark end.");
                        texts[active[id].Item1] = active[id].Item2.ToString();
                        active.Remove(id);
                    }
                    else
                        foreach (var p in active.Values)
                            p.Item2.Append(node.InnerText);
                }

                if (active.Count > 0 || texts.Keys.Count(k => k.StartsWith("LM_Message_")) != cards.Count)
                    throw new InvalidDataException("Bookmark/card count mismatch.");
                foreach (var card in cards)
                {
                    string id = card.Card.ToString("D4");
                    if (card.SourceMessage != null && card.Message != (card.EncodingRepairEnabled ? TextRepair.Fix(card.SourceMessage).Text : card.SourceMessage))
                        throw new InvalidDataException("Student response changed beyond the approved encoding repair.");
                    if (!texts.ContainsKey("LM_Name_" + id) || texts["LM_Name_" + id] != card.Name || !texts.ContainsKey("LM_Scholar_" + id) || texts["LM_Scholar_" + id] != card.Scholarship || !texts.ContainsKey("LM_Message_" + id) || !Normalize(texts["LM_Message_" + id]).Contains(Normalize(card.Message)))
                        throw new InvalidDataException("Name, scholarship or complete response mismatch in card " + card.Card);
                }

                return "PASS: package/DTD limits, internal relationships, content types, XML, unique identifiers, bookmarks, card count, exact names/scholarships and complete per-card student responses.";
            }
        }

        public static Image Preview(CardTemplate t, PhotoData photo)
        {
            if (Hash(t.Source) != t.Hash)
                throw new InvalidDataException("Template changed. Reopen preview.");
            var anchor = t.Prefix[t.FrameIndex].SelectSingleNode(".//wp:anchor[.//a:blip]", t.Ns);
            string rid = ((XmlElement)anchor.SelectSingleNode(".//a:blip", t.Ns)).GetAttribute("embed", R);
            var rel = t.Relationships.DocumentElement.ChildNodes.OfType<XmlElement>().First(e => e.GetAttribute("Id") == rid);
            var uri = new Uri(new Uri("http://a7-preview.invalid/word/document.xml"), rel.GetAttribute("Target"));
            using (var z = ZipFile.OpenRead(t.Source))
            using (var stream = z.GetEntry(Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/')).Open())
            using (var frame = Image.FromStream(stream))
            using (var memory = new MemoryStream(photo.Bytes))
            using (var image = Image.FromStream(memory))
            {
                var canvas = new Bitmap(frame.Width, frame.Height);
                try
                {
                    using (var g = Graphics.FromImage(canvas))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(frame, 0, 0, frame.Width, frame.Height);
                        var ext = (XmlElement)anchor.SelectSingleNode("wp:extent", t.Ns);
                        double fw = Double.Parse(ext.GetAttribute("cx")), fh = Double.Parse(ext.GetAttribute("cy"));
                        float width = (float)(photo.Plan.DisplayWidth / fw * frame.Width), height = (float)(photo.Plan.DisplayHeight / fh * frame.Height);
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(image, new RectangleF((frame.Width - width) / 2 + (float)(photo.Plan.X / fw * frame.Width), (frame.Height - height) / 2 + (float)(photo.Plan.Y / fh * frame.Height), width, height));
                    }

                    return canvas;
                }
                catch
                {
                    canvas.Dispose();
                    throw;
                }
            }
        }

        public static void Report(string path, string[] headers, IEnumerable<string[]> rows)
        {
            Core.RejectLink(path);
            using (var writer = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                writer.WriteLine(String.Join(",", headers.Select(ReportCell)));
                foreach (var row in rows)
                    writer.WriteLine(String.Join(",", row.Select(ReportCell)));
            }
        }

        static string ReportCell(string value)
        {
            return Quote(Regex.IsMatch(value ?? "", @"^[\t\r]|^\s*[=+@-]") ? "'" + value : value);
        }

        static string Quote(string s)
        {
            return "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
        }

        public static List<string> ContentIssues(Student s)
        {
            var issues = new List<string>(s.TextIssues);
            if (s.Words > 300)
                issues.Add("WORD_LIMIT: response exceeds 300 words; full text retained.");
            if (Regex.IsMatch(s.Message, @"\S{61,}"))
                issues.Add("LONG_UNBROKEN_TEXT");
            if (Regex.IsMatch(s.Message, @"<[^>]+>|\[(your name|name|donor|scholarship)\]", RegexOptions.IgnoreCase))
                issues.Add("MESSAGE_PLACEHOLDER_OR_MARKUP");
            if (!s.PhotoAdded)
                issues.Add("PHOTO_MISSING");
            return issues;
        }
    }
}
