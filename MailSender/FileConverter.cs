using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace MailSender
{
    class FileConverter
    {
        Word.Application WordApp;

        public FileConverter()
        {
            WordApp = new Word.Application();
            WordApp.Visible = false;
        }

        public void Convert(string openPath, string savePath)
        {
            Word.Document doc = null;
            try
            {
                doc = WordApp.Documents.Open(Path.GetFullPath(openPath));
            }
            catch
            {
                Console.WriteLine("Не смог открыть файл {0}", openPath);
                WordApp.Quit();
            }
            try
            {
                doc.SaveAs(FileName: Path.GetFullPath(savePath), FileFormat: Word.WdSaveFormat.wdFormatPDF);
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch
            {
                Console.WriteLine("Не смог сохранить файл {0}", savePath);
                WordApp.Quit();
            }

        }
        public void Free()
        {
            if (WordApp.Documents != null)
                WordApp.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            WordApp.Quit();
        }
    }
}
