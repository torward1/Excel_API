using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Excel_API
{
    internal class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string fileName) 
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File note found");
            }
        }
    }
}
