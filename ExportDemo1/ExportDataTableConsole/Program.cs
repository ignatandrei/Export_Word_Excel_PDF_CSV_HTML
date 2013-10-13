using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;


namespace ExportDataTableConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = ElectronicsDataTable.GetData();
            var exp = new ExportDataTable();
            exp.PathTemplateFolder = Path.Combine(Environment.CurrentDirectory, @"exportsTemplates\electronics");
            exp.ExportTo(dt, ExportToFormat.HTML, "a.html");
            exp.ExportTo(dt, ExportToFormat.CSV, "a.csv");
            exp.ExportTo(dt, ExportToFormat.XML, "a.xml");
            exp.ExportTo(dt, ExportToFormat.Word2003XML, "a_2003.doc");
            exp.ExportTo(dt, ExportToFormat.Excel2003XML, "a_2003.xls");
            exp.ExportTo(dt, ExportToFormat.PDFtextSharpXML, "a.pdf");
            exp.ExportTo(dt, ExportToFormat.Excel2007, "a.xlsx");
            exp.ExportTo(dt, ExportToFormat.Word2007, "a.docx");
            exp.ExportTo(dt, ExportToFormat.itextSharpXML, "a.xml");
        }
    }
}
