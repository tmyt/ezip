using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Web;
using Etier.IconHelper;

namespace ezip
{
    public class Engine
    {
        public double ImageScaling { get; set; }

        public Engine()
        {
            ImageScaling = 1.0;
        }

        public void Compress(string target, List<string> files)
        {
            var zip = ZipFile.Open(target, ZipArchiveMode.Create);
            // _rels directory
            CreateGlobalRelsXml(zip);
            CreateWorkBook(zip, files);
            // create conten types
            CreateContentTypes(zip, files);
            // create worksheets
            CreateWorksheets(zip, files);
            // create drawings
            CreateDrawings(zip, files);
            // add files
            AddFiles(zip, files);
            zip.Dispose();
        }


        void CreateGlobalRelsXml(ZipArchive zip)
        {
            var entry = zip.CreateEntry("_rels\\.rels");
            using (var stream = entry.Open())
            using (var sw = new StreamWriter(stream))
            {
                CreateRelationships(sw, new[] { new[] { "officeDocument", "workbook.xml" } });
            }
        }

        private void CreateWorkBook(ZipArchive zip, List<string> files)
        {
            // rels
            var entry = zip.CreateEntry("_rels\\workbook.xml.rels");
            using (var stream = entry.Open())
            using (var sw = new StreamWriter(stream))
            {
                CreateRelationships(sw, Enumerable
                    .Range(0, files.Count)
                    .Select(_ => new[] { "worksheet", string.Format("worksheets/sheet{0}.xml", _) }));
            }
            // workbook
            entry = zip.CreateEntry("workbook.xml");
            using (var stream = entry.Open())
            using (var sw = new StreamWriter(stream))
            {
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.WriteLine(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
                sw.WriteLine(@"	<sheets>");
                for (int i = 0; i < files.Count; i++)
                {
                    sw.WriteLine(@"		<sheet name=""{0}.{1}"" sheetId=""{2}"" r:id=""rId{0}""/>", i, Path.GetFileName(files[i]), i + 1);
                }
                sw.WriteLine(@"	</sheets>");
                sw.WriteLine(@"</workbook>");

            }
        }

        void CreateContentTypes(ZipArchive zip, List<string> files)
        {
            var entry = zip.CreateEntry("[Content_Types].xml");
            using (var stream = entry.Open())
            using (var sw = new StreamWriter(stream))
            {
                sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                sw.WriteLine(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">");
                // Defaults
                sw.WriteLine(@"	<Default Extension=""bin"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings""/>");
                sw.WriteLine(@"	<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>");
                sw.WriteLine(@"	<Default Extension=""vml"" ContentType=""application/vnd.openxmlformats-officedocument.vmlDrawing""/>");
                sw.WriteLine(@"	<Default Extension=""xml"" ContentType=""application/xml""/>");
                sw.WriteLine(@"	<Default Extension=""png"" ContentType=""image/png""/>");
                // FileTypes
                sw.WriteLine(@"	<Override PartName=""/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>");
                for (int i = 0; i < files.Count; i++)
                {
                    // Worksheet, drawings
                    sw.WriteLine(@"	<Override PartName=""/worksheets/sheet{0}.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>", i);
                    sw.WriteLine(@"	<Override PartName=""/drawings/drawing{0}.xml"" ContentType=""application/vnd.openxmlformats-officedocument.drawing+xml""/>", i);
                    // File
                    var mime = MimeMapping.GetMimeMapping(files[i]);
                    if (mime.StartsWith("image/"))
                    {
                        sw.WriteLine(@"<Override PartName=""/files/{0}"" ContentType=""{1}""/>", i + Path.GetFileName(files[i]), mime);
                    }
                    else
                    {
                        sw.WriteLine(@"<Override PartName=""/files/{0}{1}"" ContentType=""application/vnd.openxmlformats-officedocument.oleObject""/>", i, Path.GetFileName(files[i]));
                    }
                }
                sw.WriteLine(@"</Types>");
            }
        }

        void CreateWorksheets(ZipArchive zip, List<string> files)
        {
            for (int i = 0; i < files.Count; ++i)
            {
                var mime = MimeMapping.GetMimeMapping(files[i]);
                var isImage = mime.StartsWith("image/");

                // rels
                var entry = zip.CreateEntry(string.Format("worksheets/_rels/sheet{0}.xml.rels", i));
                using (var stream = entry.Open())
                using (var sw = new StreamWriter(stream))
                {
                    if (isImage)
                        CreateRelationships(sw, new[] { new[] { "drawing", string.Format("../drawings/drawing{0}.xml", i) } });
                    else
                        CreateRelationships(sw, new[]
                                                {
                                                    new[] { "drawing", string.Format("../drawings/drawing{0}.xml", i) },
                                                    new[] { "oleObject", string.Format("../files/{0}{1}", i, Path.GetFileName(files[i])) }
                                                });
                }

                // content
                entry = zip.CreateEntry(string.Format("worksheets/sheet{0}.xml", i));
                using (var stream = entry.Open())
                using (var sw = new StreamWriter(stream))
                {
                    sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                    sw.WriteLine(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""><sheetData/>");
                    sw.WriteLine(@"<drawing r:id=""rId0""/>");
                    if (!isImage)
                    {
                        sw.WriteLine(@"	<oleObjects><oleObject progId=""パッケージャー シェル オブジェクト"" shapeId=""1025"" r:id=""rId1""/></oleObjects>");
                    }
                    sw.WriteLine("</worksheet>");
                }
            }
        }

        void CreateDrawings(ZipArchive zip, List<string> files)
        {
            for (int i = 0; i < files.Count; ++i)
            {
                var mime = MimeMapping.GetMimeMapping(files[i]);
                var isImage = mime.StartsWith("image/");

                if (!isImage)
                {
                    CreateThumbnail(zip, string.Format("image{0}.png", i), files[i]);
                }

                // rels
                var entry = zip.CreateEntry(string.Format("drawings/_rels/drawing{0}.xml.rels", i));
                using (var stream = entry.Open())
                using (var sw = new StreamWriter(stream))
                {
                    var fileName = string.Format("image{0}.png", i);
                    if (isImage) fileName = string.Format("../files/{0}{1}", i, Path.GetFileName(files[i]));
                    CreateRelationships(sw, new[] { new[] { "image", fileName } });
                }

                // content
                var w = 180;
                var h = 70;
                if (isImage)
                {
                    var bitmap = new Bitmap(files[i]);
                    w = (int)(bitmap.Width * ImageScaling);
                    h = (int)(bitmap.Height * ImageScaling);
                }
                w = w * 914400 / 96;
                h = h * 914400 / 96;

                entry = zip.CreateEntry(string.Format("drawings/drawing{0}.xml", i));
                using (var stream = entry.Open())
                using (var sw = new StreamWriter(stream))
                {
                    sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
                    sw.WriteLine(@"<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">");
                    sw.WriteLine(@"	<xdr:absoluteAnchor>");
                    sw.WriteLine(@"		<xdr:pos x=""91440"" y=""91440""/><xdr:ext cx=""{0}"" cy=""{1}"" />", w, h);
                    sw.WriteLine(@"		<xdr:pic>");
                    sw.WriteLine(@"			<xdr:nvPicPr><xdr:cNvPr id=""2"" name=""図 1""/><xdr:cNvPicPr><a:picLocks noChangeAspect=""1""/></xdr:cNvPicPr></xdr:nvPicPr>");
                    sw.WriteLine(@"			<xdr:blipFill><a:blip xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:embed=""rId0"" />");
                    sw.WriteLine(@"<a:stretch/>");
                    sw.WriteLine(@"</xdr:blipFill>");
                    sw.WriteLine(@"			<xdr:spPr><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom></xdr:spPr>");
                    sw.WriteLine(@"		</xdr:pic>");
                    sw.WriteLine(@"		<xdr:clientData/>");
                    sw.WriteLine(@"	</xdr:absoluteAnchor>");
                    sw.WriteLine(@"</xdr:wsDr>");
                }
            }
        }

        void CreateRelationships(StreamWriter sw, List<Relationship> relationships)
        {
            sw.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            sw.WriteLine(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">");
            for (int i = 0; i < relationships.Count; ++i)
            {
                sw.WriteLine(@"<Relationship Id=""rId{0}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/{1}"" Target=""{2}""/>",
                    i, relationships[i].Type, relationships[i].Target);
            }
            sw.WriteLine(@"</Relationships>");
        }

        void CreateRelationships(StreamWriter sw, IEnumerable<string[]> relationships)
        {
            CreateRelationships(sw, relationships.Select(a => new Relationship { Type = a[0], Target = a[1] }).ToList());
        }

        void CreateThumbnail(ZipArchive zip, string fileName, string filePath)
        {
            var w = 180;
            var h = 70;
            var bmp = new Bitmap(w, h);
            var icon = IconReader.GetFileIcon(filePath, IconReader.IconSize.Large, false);
            var g = Graphics.FromImage(bmp);
            var font = new Font("Segoe UI", 12);
            var s = Path.GetFileName(filePath);
            g.FillRectangle(Brushes.White, 0, 0, w, h);
            g.DrawRectangle(Pens.Black, 0, 0, w - 1, h - 1);
            g.DrawIcon(icon, (w - icon.Width) / 2, 10);
            var r = g.MeasureString(s, font);
            g.DrawString(s, font, Brushes.Black, new PointF((w - r.Width) / 2, 40));
            g.Flush();
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                var entry = zip.CreateEntry("drawings/" + fileName);
                using (var stream = entry.Open())
                {
                    ms.Seek(0, SeekOrigin.Begin);
                    ms.CopyTo(stream);
                }
            }
        }

        private void AddFiles(ZipArchive zip, List<string> files)
        {
            for (int i = 0; i < files.Count; ++i)
            {
                zip.CreateEntryFromFile(files[i], string.Format("files/{0}{1}", i, Path.GetFileName(files[i])));
            }
        }
    }

    internal class Relationship
    {
        public string Type { get; set; }
        public string Target { get; set; }
    }
}
