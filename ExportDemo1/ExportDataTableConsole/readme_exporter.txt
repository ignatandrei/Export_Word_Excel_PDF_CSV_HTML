To export  a List<Electronics> into variouse formats, just make an exportsTemplates\electronics folder and put those lines:


List<Electronics> list = Electronics.GetData();
                ExportList<Electronics> exp = new ExportList<Electronics>();
                exp.PathTemplateFolder=Path.Combine(Environment.CurrentDirectory,  @"exportsTemplates\electronics");                
                exp.ExportTo(list, ExportToFormat.HTML, "a.html");
                exp.ExportTo(list, ExportToFormat.CSV, "a.csv");
                exp.ExportTo(list, ExportToFormat.XML, "a.xml");
                exp.ExportTo(list, ExportToFormat.Word2003XML, "a_2003.doc");
                exp.ExportTo(list, ExportToFormat.Excel2003XML, "a_2003.xls");
                exp.ExportTo(list, ExportToFormat.Excel2007, "a.xlsx");
                exp.ExportTo(list, ExportToFormat.Word2007, "a.docx");
                exp.ExportTo(list, ExportToFormat.itextSharpXML, "a.xml");
                exp.ExportTo(list, ExportToFormat.PDFtextSharpXML, "a.pdf");
                
                See video at http://youtu.be/DHNRV9hzG_Y
                
For Asp.net MVC  you can export with

List<Electronics> list = Electronics.GetData();
            ExportList<Electronics> exp = new ExportList<Electronics>();
            exp.PathTemplateFolder = Server.MapPath("~/ExportTemplates/electronics");

            string filePathExport = Server.MapPath("~/exports/a" + ExportBase.GetFileExtension((ExportToFormat)id));
            exp.ExportTo(list, (ExportToFormat)id, filePathExport);
            
See video at http://www.youtube.com/edit?video_id=2CBdn6ru47M&o=U&ns=1            
            

You can export also a DataTable


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
            
            
            
            
            
            
You can modify also the generated CSHTML template to improve your exported data.
                
                
                