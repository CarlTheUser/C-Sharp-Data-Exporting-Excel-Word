using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExporting
{
    public class FieldDisplayAttribute : Attribute
    {
        public string Title { get; }
        public string Format { get; }
        public bool IsIncluded { get; }

        public FieldDisplayAttribute(string title, string format = "")
        {
            Title = title;
            Format = format;
            IsIncluded = true;
        }

        public FieldDisplayAttribute(bool include) => IsIncluded = include;
    }
}
