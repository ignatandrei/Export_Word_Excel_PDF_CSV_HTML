using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;

namespace ExporterObjects
{
    public class ExportDataTable : ExportBase
    {
        public void ExportTo(DataTable dt, ExportToFormat exp, string fileName)
        {
            if (string.IsNullOrEmpty(dt.TableName))
                throw new ArgumentException("Please provide a table name for datatable ");

            var tableName = dt.TableName;
            var path1 = Path.GetDirectoryName(fileName);
            if (string.IsNullOrEmpty(path1))
                path1 = Environment.CurrentDirectory;

            var path = Path.Combine(path1, tableName + ".dll");
            Type type;

            if (!File.Exists(path))
            {
                var provider = CodeDomProvider.CreateProvider("CSharp");
                var options = new CompilerParameters();
                options.ReferencedAssemblies.Add("System.dll");
                options.ReferencedAssemblies.Add("System.Data.dll");
                options.ReferencedAssemblies.Add("System.XML.dll");
                options.GenerateExecutable = false;
                options.GenerateInMemory = false;
                options.OutputAssembly = path;

                var stringBuilder = new StringBuilder();
                stringBuilder.Append("using System;");
                stringBuilder.Append("using System.Data;");
                stringBuilder.Append("using System.Collections.Generic;");
                stringBuilder.AppendFormat("namespace Export{0} {{", tableName);
                stringBuilder.AppendFormat("public class {0} {{", tableName);
                foreach (var obj in dt.Columns)
                {
                    var dataColumn = obj as DataColumn;
                    stringBuilder.AppendFormat("public {0} {1} {{get;set;}}", dataColumn.DataType.FullName, dataColumn.ColumnName);
                }

                stringBuilder.AppendFormat("public static {0} From(DataRow dr) {{", tableName);
                stringBuilder.AppendFormat("{0} ret=new {0}();", tableName);
                foreach (var obj in dt.Columns)
                {
                    var dataColumn = obj as DataColumn;
                    stringBuilder.AppendFormat("ret.{0}=  (dr[{1}] == null || dr[{1}] == DBNull.Value)?default({2}):({2}) dr[{1}];", dataColumn.ColumnName, dataColumn.Ordinal, dataColumn.DataType.FullName);
                }

                stringBuilder.Append("return ret;");
                stringBuilder.Append("}");
                stringBuilder.AppendFormat("public static List<{0}> FromDataTable(DataTable dt) {{", tableName);
                stringBuilder.AppendFormat("List<{0}> ret=new List<{0}>();", tableName);
                stringBuilder.Append("foreach(var row in dt.Rows){");
                stringBuilder.Append("ret.Add(From(row as DataRow));");
                stringBuilder.Append("}");
                stringBuilder.Append("return ret;");
                stringBuilder.Append("}");
                stringBuilder.Append("}");
                stringBuilder.Append("}");

                var compilerResults = provider.CompileAssemblyFromSource(options, stringBuilder.ToString());
                if (compilerResults.Errors.Count > 0)
                    throw new ArgumentException(compilerResults.Errors[0].ErrorText);
                type = compilerResults.CompiledAssembly.GetType(string.Format("Export{0}.{0}", tableName), true, false);
            }
            else
                type = Assembly.LoadFile(path).GetType(string.Format("Export{0}.{0}", tableName), true, false);

            ExportTo(type.GetMethod("FromDataTable", BindingFlags.Static | BindingFlags.Public).Invoke(null, new object[] { dt }) as IEnumerable, type, exp, fileName);
        }
    }
}
