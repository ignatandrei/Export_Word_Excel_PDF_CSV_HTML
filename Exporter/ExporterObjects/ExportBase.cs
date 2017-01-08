using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Westwind.RazorHosting;

namespace ExporterObjects
{
    public abstract class ExportBase
    {
        private Dictionary<string, Dictionary<string, MethodInfo>> _methods;
        private Dictionary<string, List<string>> _nameprops;

        public string TemplateFileName { get; set; }
        public string PathTemplateFolder { get; set; }

        private void FillNameprops(Type type)
        {
            if (_methods == null)
            {
                _methods = new Dictionary<string, Dictionary<string, MethodInfo>>();
                _nameprops = new Dictionary<string, List<string>>();
            }

            if (_methods.ContainsKey(type.FullName))
                return;

            var dictionary = new Dictionary<string, MethodInfo>();
            var list = new List<string>();
            foreach (var propertyInfo in type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                if (!(propertyInfo.GetMethod == null) && propertyInfo.GetMethod.GetParameters().Length <= 0)
                {
                    list.Add(propertyInfo.Name);
                    dictionary.Add(propertyInfo.Name, propertyInfo.GetGetMethod());
                }
            }

            _methods.Add(type.FullName, dictionary);
            _nameprops.Add(type.FullName, list);
        }

        protected void WriteitextSharpXMLTemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<row>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<cell><phrase font='Times New Roman' size='8'>@Model.{0}</phrase></cell>", str);

            stringBuilder1.AppendLine("</row>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<itext author='Andrei Ignat' title='Collection'>\r\n                <chapter numberdepth='0'>\r\n                <newline />\r\n                <section numberdepth='0'>\r\n                <table width='100%'  cellspacing='0' cellpadding='2' columns='" + (object)_nameprops[type.FullName].Count + "' grayfill='0.90'>                \r\n            ");
            stringBuilder2.AppendLine("<row>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat("<cell><phrase font='Arial' size='12' style='bold'>{0}</phrase></cell>", str);

            stringBuilder2.AppendLine("</row>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("</table></section></chapter></itext>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteWord2007TemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<w:tr>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<w:tc><w:p><w:r><w:t>@Model.{0}</w:t></w:r></w:p></w:tc>", str);

            stringBuilder1.AppendLine("</w:tr>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:tbl>");
            stringBuilder2.AppendLine("<w:tr>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat("<w:tc><w:p><w:r><w:rPr><w:b w:val='on'/><w:t>{0}</w:t></w:rPr></w:r></w:p></w:tc>", str);

            stringBuilder2.AppendLine("</w:tr>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("\t\t</w:tbl></w:body></w:document>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteExcel2007TemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<row>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<c t='inlineStr'><is><t>@Model.{0}</t></is></c>", str);

            stringBuilder1.AppendLine("</row>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<?xml version='1.0' encoding='UTF-8' standalone='yes' ?>\r\n<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>\r\n\t<sheetData>\r\n   ");
            stringBuilder2.AppendLine("<row>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat("<c t='inlineStr'><is><t>{0}</t></is></c>", str);

            stringBuilder2.AppendLine("</row>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("\t</sheetData></worksheet>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteExcel2003TemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<Row>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<Cell><Data ss:Type='String'>@Model.{0}</Data></Cell>", str);
            stringBuilder1.AppendLine("</Row>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<?xml version='1.0'?>\r\n<?mso-application progid='Excel.Sheet'?>\r\n<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet'\r\n xmlns:o='urn:schemas-microsoft-com:office:office'\r\n xmlns:x='urn:schemas-microsoft-com:office:excel'\r\n xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'\r\n xmlns:html='http://www.w3.org/TR/REC-html40'>\r\n <DocumentProperties xmlns='urn:schemas-microsoft-com:office:office'>\r\n  <Author>Andrei Ignat</Author>\r\n  <LastAuthor>Andrei Ignat</LastAuthor>\r\n  <Created>@DateTime.Now.ToString(\"yyyy-MM-ddTHH:mm:ssZ\")</Created>\r\n  <Company>AOM</Company>\r\n  <Version>11.9999</Version>\r\n </DocumentProperties>\r\n <ExcelWorkbook xmlns='urn:schemas-microsoft-com:office:excel'>\r\n  <WindowHeight>11760</WindowHeight>\r\n  <WindowWidth>15195</WindowWidth>\r\n  <WindowTopX>480</WindowTopX>\r\n  <WindowTopY>90</WindowTopY>\r\n  <ProtectStructure>False</ProtectStructure>\r\n  <ProtectWindows>False</ProtectWindows>\r\n </ExcelWorkbook>\r\n <Styles>\r\n  <Style ss:ID='Default' ss:Name='Normal'>\r\n   <Alignment ss:Vertical='Bottom'/>\r\n   <Borders/>\r\n   <Font/>\r\n   <Interior/>\r\n   <NumberFormat/>\r\n   <Protection/>\r\n  </Style>\r\n  <Style ss:ID='s21'>\r\n   <Font x:Family='Swiss' ss:Bold='1'/>\r\n  </Style>\r\n </Styles>\r\n <Worksheet ss:Name='Sheet1'>\r\n  <Table  x:FullColumns='1'\r\n   x:FullRows='1'>\r\n   ");
            stringBuilder2.AppendLine("<Row>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat("<Cell ss:StyleID='s21'><Data ss:Type='String'>{0}</Data></Cell>", str);

            stringBuilder2.AppendLine("</Row>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("</Table>\r\n  <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>\r\n   <Print>\r\n    <ValidPrinterInfo/>\r\n    <PaperSizeIndex>9</PaperSizeIndex>\r\n    <HorizontalResolution>600</HorizontalResolution>\r\n    <VerticalResolution>600</VerticalResolution>\r\n   </Print>\r\n   <Selected/>\r\n   <Panes>\r\n    <Pane>\r\n     <Number>3</Number>\r\n     <RangeSelection>R1C1:R1C1</RangeSelection>\r\n    </Pane>\r\n   </Panes>\r\n   <ProtectObjects>False</ProtectObjects>\r\n   <ProtectScenarios>False</ProtectScenarios>\r\n  </WorksheetOptions>\r\n </Worksheet>\r\n</Workbook>\r\n");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteWord2003TemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<w:tr>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<w:tc><w:p><w:r><w:t>@Model.{0}</w:t></w:r></w:p></w:tc>", str);

            stringBuilder1.AppendLine("</w:tr>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<?xml version='1.0'?><w:wordDocument xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:body><w:tbl>");
            stringBuilder2.AppendLine("<w:tr>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat("<w:tc><w:p><w:r><w:rPr><w:b w:val='on'/><w:t>{0}</w:t></w:rPr></w:r></w:p></w:tc>", str);

            stringBuilder2.AppendLine("</w:tr>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("</w:tbl></w:body></w:wordDocument>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteXMLTemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<item>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<{0}>@Model.{0}</{0}>", str);

            stringBuilder1.AppendLine("</item>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<root><dictionary>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat("<prop>{0}</prop>", str);

            stringBuilder2.AppendLine("</dictionary>");
            stringBuilder2.AppendLine("<values>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("</values>");
            stringBuilder2.AppendLine("</root>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteCSVTemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            var flag = true;
            foreach (var str1 in _nameprops[type.FullName])
            {
                var str2 = flag ? "" : ",";
                stringBuilder1.AppendFormat("{0}\"@Model.{1}\"", str2, str1);
                flag = false;
            }
            stringBuilder1.AppendLine("");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            flag = true;
            foreach (var str1 in _nameprops[type.FullName])
            {
                var str2 = flag ? "" : ",";
                stringBuilder2.AppendFormat("{0}\"{1}\"", str2, str1);
                flag = false;
            }
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        protected void WriteHtmlTemplateForIEnumerable(string templateFile, Type type)
        {
            FillNameprops(type);

            var path = Path.GetFileNameWithoutExtension(templateFile) + "_item.cshtml";
            var stringBuilder1 = new StringBuilder("@inherits RazorTemplateFolderHost<" + type.FullName + ">");
            stringBuilder1.AppendLine("");
            stringBuilder1.AppendLine("<tr>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder1.AppendFormat("<td>@Model.{0}</td>", str);

            stringBuilder1.AppendLine("</tr>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, path), stringBuilder1.ToString());

            var stringBuilder2 = new StringBuilder("@inherits RazorTemplateFolderHost<IEnumerable<" + type.FullName + ">>");
            stringBuilder2.AppendLine("");
            stringBuilder2.AppendLine("<html><head><title>Export</title></head><body>");
            stringBuilder2.AppendLine("<table><tr>");
            foreach (var str in _nameprops[type.FullName])
                stringBuilder2.AppendFormat(" <td>{0}</td>", str);

            stringBuilder2.AppendLine("</tr>");
            stringBuilder2.AppendLine("@foreach(var item in Model){");
            stringBuilder2.AppendLine("@RenderPartial(\"~/" + path + "\",item);");
            stringBuilder2.AppendLine("}");
            stringBuilder2.AppendLine("</table>");
            stringBuilder2.AppendLine("</body>");
            stringBuilder2.AppendLine("Generated on @DateTime.Now.ToString()");
            stringBuilder2.AppendLine("</html>");
            File.WriteAllText(Path.Combine(PathTemplateFolder, templateFile), stringBuilder2.ToString());
        }

        private RazorFolderHostContainer CreateRazorFolderHostContainer(Type type, [In] string str)
        {
            var folderHostContainer = new RazorFolderHostContainer
            {
                TemplatePath = str,
                BaseBinaryFolder = Environment.CurrentDirectory,
                UseAppDomain = false,
                Configuration = {CompileToMemory = true}
            };

            folderHostContainer.AddAssemblyFromType(type);
            folderHostContainer.Start();

            return folderHostContainer;
        }

        protected void CreateFilePDF(string itextSharpXml, string fileName)
        {
            var str = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetTempFileName()));
            File.WriteAllText(str + ".xml", itextSharpXml);
            var document = new Document();
            PdfWriter.GetInstance(document, new FileStream(fileName, FileMode.Create));
            XmlParser.Parse(document, new XmlTextReader(str + ".xml")
            {
                WhitespaceHandling = WhitespaceHandling.None
            });
            document.Close();
        }

        protected void ExportTo(IEnumerable col, Type type, ExportToFormat exp, string fileName)
        {
            var str1 = TemplateFileName;
            if (string.IsNullOrWhiteSpace(str1))
                str1 = type.Name + "_" + exp + ".cshtml";
            if (!File.Exists(Path.Combine(PathTemplateFolder, str1)))
            {
                switch (exp)
                {
                    case ExportToFormat.Word2003XML:
                        WriteWord2003TemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.Excel2003XML:
                        WriteExcel2003TemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.CSV:
                        WriteCSVTemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.HTML:
                        WriteHtmlTemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.XML:
                        WriteXMLTemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.itextSharpXML:
                    case ExportToFormat.PDFtextSharpXML:
                        WriteitextSharpXMLTemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.Word2007:
                        WriteWord2007TemplateForIEnumerable(str1, type);
                        break;
                    case ExportToFormat.Excel2007:
                        WriteExcel2007TemplateForIEnumerable(str1, type);
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }

            var list = CreateInstance(type);
            foreach (var obj in col)
                list.Add(obj);

            var folderHostContainer = CreateRazorFolderHostContainer(type, PathTemplateFolder);
            var str2 = folderHostContainer.RenderTemplate(str1, list);
            if (!string.IsNullOrWhiteSpace(folderHostContainer.ErrorMessage))
                throw new Exception(folderHostContainer.ErrorMessage);

            switch (exp)
            {
                case ExportToFormat.PDFtextSharpXML:
                    CreateFilePDF(str2, fileName);
                    break;
                case ExportToFormat.Word2007:
                    CreateWord2007(str2, fileName);
                    break;
                case ExportToFormat.Excel2007:
                    CreateExcel2007(str2, fileName);
                    break;
                default:
                    File.WriteAllText(fileName, str2);
                    break;
            }
        }

        private IList CreateInstance(Type type)
        {
            return Activator.CreateInstance(typeof(List<>).MakeGenericType(type)) as IList;
        }

        protected void CreateWord2007(string text, string fileName)
        {
            using (var wordprocessingDocument = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                WriteToPart(wordprocessingDocument.AddMainDocumentPart(), text);
                wordprocessingDocument.Close();
            }
        }

        protected void CreateExcel2007(string sheet1Text, string destinationFile)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(destinationFile, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                WriteToPart(worksheetPart, sheet1Text);
                WriteToPart(workbookPart, string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"{1}\" sheetId=\"1\" r:id=\"{0}\" /></sheets></workbook>", (object)workbookPart.GetIdOfPart((OpenXmlPart)worksheetPart), (object)"Sheet1"));
                spreadsheetDocument.Close();
            }
        }

        protected void WriteToPart(OpenXmlPart oxp, string text)
        {
            using (var stream = oxp.GetStream())
            {
                var bytes = new UTF8Encoding().GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
            }
        }

        public static string GetFileExtension(ExportToFormat e)
        {
            switch (e)
            {
                case ExportToFormat.Word2003XML:
                    return ".doc";
                case ExportToFormat.Excel2003XML:
                    return ".xls";
                case ExportToFormat.CSV:
                    return ".csv";
                case ExportToFormat.HTML:
                    return ".html";
                case ExportToFormat.XML:
                    return ".xml";
                case ExportToFormat.itextSharpXML:
                    return ".xml";
                case ExportToFormat.PDFtextSharpXML:
                    return ".pdf";
                case ExportToFormat.Word2007:
                    return ".docx";
                case ExportToFormat.Excel2007:
                    return ".xlsx";
                default:
                    throw new ArgumentException("not found : " + (object)e);
            }
        }
    }
}
