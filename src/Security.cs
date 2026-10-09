using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Xml;

namespace LetterMerger
{
    public static class Security
    {
        static readonly HashSet<string> TemplateParts = new HashSet<string>(StringComparer.Ordinal)
        {
            "[Content_Types].xml", "_rels/.rels", "word/document.xml",
            "word/_rels/document.xml.rels", "word/footnotes.xml", "word/endnotes.xml",
            "word/theme/theme1.xml", "word/settings.xml", "word/styles.xml",
            "word/webSettings.xml", "word/fontTable.xml", "docProps/core.xml", "docProps/app.xml"
        };

        static readonly HashSet<string> PlaceholderText = new HashSet<string>(StringComparer.Ordinal)
        {
            "", "IMAGE", "\u00abName\u00bb", "Recipient of the ", "\u00abPortfolio_Name\u00bb",
            "Dear Donor,", "\u00abPlease_draft_a_Thank_You_letter_no_mo\u00bb", "Sincerely,"
        };

        static readonly HashSet<string> RelationshipTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            Core.R + "/officeDocument", Core.R + "/image", Core.R + "/theme",
            Core.R + "/webSettings", Core.R + "/fontTable", Core.R + "/settings",
            Core.R + "/styles", Core.R + "/footnotes", Core.R + "/endnotes",
            Core.R + "/extended-properties",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        };

        public static void CheckTemplate(ZipArchive archive)
        {
            var media = archive.Entries.Where(e => e.FullName.StartsWith("word/media/", StringComparison.Ordinal)).ToList();
            if (media.Count != 1 || !Regex.IsMatch(media[0].FullName, @"^word/media/[A-Za-z0-9_-]+\.(jpeg|jpg|png)$"))
                throw new InvalidDataException("Use the original blank A7 template with its single frame image.");
            foreach (var part in archive.Entries)
            {
                if (!TemplateParts.Contains(part.FullName) && part != media[0])
                    throw new InvalidDataException("The template contains unsupported extra parts. Use the original blank A7 template.");
                if (!part.FullName.EndsWith(".xml", StringComparison.Ordinal))
                    continue;
                var document = Core.ReadXml(archive, part.FullName);
                var ns = Core.Names(document);
                if (document.SelectNodes("//w:ins|//w:del|//w:moveFrom|//w:moveTo|//w:commentRangeStart|//w:commentRangeEnd|//w:commentReference|//w:customXml|//w:dataBinding|//w:mailMerge|//w:subDoc|//w:object|//w:control", ns).Count > 0)
                    throw new InvalidDataException("Tracked changes, comments, data connections and embedded objects are unsupported in templates.");
                foreach (XmlElement hidden in document.SelectNodes("//w:vanish|//w:webHidden", ns))
                    if (!new[] { "0", "false", "off" }.Contains(hidden.GetAttribute("val", Core.W)))
                        throw new InvalidDataException("Hidden template text is unsupported.");
                if (part.FullName == "word/document.xml")
                {
                    foreach (XmlNode text in document.SelectNodes("//w:t", ns))
                        if (!PlaceholderText.Contains(text.InnerText))
                            throw new InvalidDataException("The template contains filled-in or unexpected text. Use the original blank A7 template.");
                    foreach (XmlNode instruction in document.SelectNodes("//w:instrText|//w:fldSimple/@w:instr", ns))
                        if (!Regex.IsMatch(instruction.InnerText, @"^\s*MERGEFIELD\s+(Name|Portfolio_Name|Please_draft_a_Thank_You_letter[A-Za-z0-9_]*)\s*(\\\*\s*MERGEFORMAT\s*)?$"))
                            throw new InvalidDataException("The template contains an unsupported Word field.");
                }
                else if (document.SelectNodes("//w:t|//w:instrText|//w:fldSimple", ns).Count > 0)
                    throw new InvalidDataException("Text outside the main blank template is unsupported.");
            }
            CheckRelationships(archive);
            using (var stream = media[0].Open())
            using (var digest = System.Security.Cryptography.SHA256.Create())
                if (BitConverter.ToString(digest.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() != "bbb9d9579cf7f0fb8bdf0bdd1ba16da60b5f1b8a2c70723aeb7fc83b717ab45f")
                    throw new InvalidDataException("The frame image differs from the approved blank design. Use the original unchanged A7 template.");
        }

        public static void CheckRelationships(ZipArchive archive)
        {
            var parts = new HashSet<string>(archive.Entries.Select(e => e.FullName), StringComparer.Ordinal);
            foreach (var part in archive.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
            {
                var document = Core.ReadXml(archive, part.FullName);
                string source = "";
                if (part.FullName != "_rels/.rels")
                {
                    var match = Regex.Match(part.FullName, @"^(.*)/_rels/([^/]+)\.rels$");
                    if (!match.Success)
                        throw new InvalidDataException("Invalid template relationship part.");
                    source = match.Groups[1].Value + "/" + match.Groups[2].Value;
                }
                var ids = new HashSet<string>(StringComparer.Ordinal);
                foreach (XmlElement relation in document.DocumentElement.ChildNodes.OfType<XmlElement>())
                {
                    if (String.IsNullOrEmpty(relation.GetAttribute("Id")) || !ids.Add(relation.GetAttribute("Id")))
                        throw new InvalidDataException("Repeated or missing template relationship ID.");
                    if (String.Equals(relation.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase) || !RelationshipTypes.Contains(relation.GetAttribute("Type")))
                        throw new InvalidDataException("External or unsupported template relationships are blocked.");
                    var uri = new Uri(new Uri("http://a7-package.invalid/" + source), relation.GetAttribute("Target"));
                    string target = Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/');
                    if (uri.Scheme != "http" || uri.Host != "a7-package.invalid" || !parts.Contains(target))
                        throw new InvalidDataException("An internal template relationship target is invalid.");
                }
            }
        }

        public static Dictionary<string, byte[]> OutputMetadata()
        {
            return new Dictionary<string, byte[]>
            {
                { "docProps/core.xml", Encoding.UTF8.GetBytes("<?xml version='1.0' encoding='UTF-8'?><cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'/>") },
                { "docProps/app.xml", Encoding.UTF8.GetBytes("<?xml version='1.0' encoding='UTF-8'?><Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'><Application>Letter Merger</Application></Properties>") }
            };
        }
    }
}
