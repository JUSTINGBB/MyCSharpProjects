using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace wordTableToExcel.component.openxmlOpearte
{
    class WordOperate
    {
        public static void ReadWord(string fileName)
        {
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(fileName, true))
            {
                // Insert other code here.
            }
        }
        
    }
}
