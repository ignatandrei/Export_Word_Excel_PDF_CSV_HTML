using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExporterObjects;



namespace ExportDemo1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Electronics> list = Electronics.GetData();
            ExportList<Electronics> exp = new ExportList<Electronics>();
            exp.PathTemplateFolder = Path.Combine(Environment.CurrentDirectory, "templates/electronics");
            
            exp.ExportTo(list, ExportToFormat.HTML, "a.html");
            exp.ExportTo(list, ExportToFormat.CSV, "a.csv");
            exp.ExportTo(list, ExportToFormat.XML, "a.xml");
            exp.ExportTo(list, ExportToFormat.Word2003XML, "a_2003.doc");
            exp.ExportTo(list, ExportToFormat.Excel2003XML, "a_2003.xls");
            exp.ExportTo(list, ExportToFormat.Excel2007, "a.xlsx");
            exp.ExportTo(list, ExportToFormat.Word2007, "a.docx");
            exp.ExportTo(list, ExportToFormat.itextSharpXML, "a.xml");
            exp.ExportTo(list, ExportToFormat.PDFtextSharpXML, "a.pdf");
                
            
            
           
        }
    }
}
