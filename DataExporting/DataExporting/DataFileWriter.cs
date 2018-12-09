using System;
using System.Runtime.InteropServices;

namespace DataExporting
{
    public abstract class DataFileWriter : IDataWriter, IDisposable
    {
        public string TargetPath { get; set; } = string.Empty;
        
        public void Write()
        {
            if (TargetPath == string.Empty) throw new Exception("No File target specified.");
            Write(TargetPath);
        }

        public abstract void Write(string targetPath);

        public abstract void Dispose();
        
        protected static void Deallocate(params object[] objects)
        {
            int size = objects.Length;
            try
            {
                for(int i = size-1; i != 0; --i) while (Marshal.ReleaseComObject(objects[i]) > 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                for (int i = size - 1; i != 0; --i) objects[i] = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
