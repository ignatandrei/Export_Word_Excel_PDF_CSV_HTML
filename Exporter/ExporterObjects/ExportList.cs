using System.Collections.Generic;

namespace ExporterObjects
{
    public class ExportList<T> : ExportBase
    {
        public void ExportTo(IEnumerable<T> col, ExportToFormat exp, string fileName)
        {
            ExportTo(col, typeof (T), exp, fileName);
        }
    }
}
