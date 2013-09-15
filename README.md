Export_Word_Excel_PDF_CSV_HTML
==============================

this project shows how to export to word / excel / pdf / csv / html

This is so simple:

            List<Electronics> list = Electronics.GetData(); // obtain data
            ExportList<Electronics> exp = new ExportList<Electronics>(); // make new export class
            exp.PathTemplateFolder = Path.Combine(Environment.CurrentDirectory, "templates/electronics"); // here are the templates - auto-generated if not exists
            //export now in what formats you want 
            exp.ExportTo(list, ExportToFormat.HTML, "a.html");
            exp.ExportTo(list, ExportToFormat.CSV, "a.csv");
            exp.ExportTo(list, ExportToFormat.XML, "a.xml");
            exp.ExportTo(list, ExportToFormat.Word2003XML, "a_2003.doc");
            exp.ExportTo(list, ExportToFormat.Excel2003XML, "a_2003.xls");
            exp.ExportTo(list, ExportToFormat.Excel2007, "a.xlsx");
            exp.ExportTo(list, ExportToFormat.Word2007, "a.docx");
            exp.ExportTo(list, ExportToFormat.itextSharpXML, "a.xml");
            exp.ExportTo(list, ExportToFormat.PDFtextSharpXML, "a.pdf");
            
            
The Nuget package is at http://www.nuget.org/packages/Exporter/
YouTube demo at http://youtu.be/2CBdn6ru47M
More details at http://msprogrammer.serviciipeweb.ro/2013/09/16/export-to-word-excel-pdf-csv-html/
