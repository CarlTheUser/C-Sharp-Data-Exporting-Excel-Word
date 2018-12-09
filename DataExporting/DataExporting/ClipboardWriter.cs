using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace DataExporting
{
    class ClipboardWriter : IDataWriter
    {
        public object CurrentData { get; set; }

        public void Write()
        {
            Clipboard.SetDataObject(CurrentData);
        }

        public void Write(object data)
        {
            CurrentData = data;
            Write();
        }
    }
}
