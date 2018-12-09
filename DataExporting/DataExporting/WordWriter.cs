using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;

namespace DataExporting
{
    public class WordWriter : DataFileWriter
    {
        private Word.Application _application;

        private Word.Document _document;

        private readonly object NoValue = Missing.Value;

        private readonly object DocumentEnd = "\\endofdoc";

        private const int DEFAULT_SPACING = 6;

        public WordWriter()
        {
            _application = new Word.Application();
            _document = _application.Documents.Add(ref NoValue, ref NoValue, ref NoValue, ref NoValue);
            


        }

        public WordWriter AppendParagraph(string text)
        {
            //_document.Content.SetRange(0, 0);
            _document.Content.Text += text;
            return this;
        }

        public WordWriter AppendTable<TData>(IEnumerable<TData> source) where TData : class
        {



            return this;
        }

        public WordWriter SkipLines(int lineCount = 1)
        {

            return this;
        }

        public WordWriter InsertTab()
        {

            return this;
        }

        public WordWriter NewPage()
        {
            return this;
        }

        public WordWriter BreakPage()
        {

            return this;
        }

        public WordWriter AppendImage(string imageFilePath)
        {
            return this;
        }

        public override void Write(string targetPath)
        {
            
        }

        public override void Dispose()
        {
            Deallocate(_application, _document);   
        }

        public enum BreakType
        {
            Page,
            Column,
            TextWrapping,
            NextPage,
            Continuous,
            Even,
            Odd
        }
    }
}

