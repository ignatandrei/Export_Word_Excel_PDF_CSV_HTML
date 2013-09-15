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
                
                
                You can modify also the generated CSHTML template