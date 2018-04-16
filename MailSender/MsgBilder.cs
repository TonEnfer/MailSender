using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using TemplateEngine.Docx;
using System.Linq.Expressions;

namespace MailSender
{
    public static class MsgFileBilder
    {
        public struct TextParameters
        {
            public string name;
            public string value;
            public TextParameters(string name, string value)
            {
                this.value = value;
                this.name = name;
            }
        }

        public static void Build(List<TextParameters> param, string templatePath)
        {
            string saveDir = Path.GetDirectoryName(templatePath) + "\\Generated";
            if (!Directory.Exists(saveDir))
            {
                Directory.CreateDirectory(saveDir);
            }
            if (Path.GetExtension(templatePath) == ".txt")
            {
                string text = Encoding.Default.GetString(File.ReadAllBytes(templatePath));
                foreach (var ppp in param)
                {
                    if (text.Contains("{" + ppp.name + "}"))
                    {
                        text = text.Replace("{"+ ppp.name + "}",ppp.value);
                    }
                }
                
                
                string newFilePath = saveDir + "\\" + Path.GetFileNameWithoutExtension(templatePath) + "-" +
                    (from n in param where n.name == "MsgNumber" select n.value).First() + ".txt";
                
                File.WriteAllText(newFilePath, text);
            }
            else if (Path.GetExtension(templatePath) == ".docx")
            {
                var valuesToFill = new Content();
                foreach (var ppp in param)
                {
                    FieldContent f = new FieldContent(ppp.name, ppp.value);
                    valuesToFill.Fields.Add(f);
                }
                string newFilePath = saveDir + "\\" + Path.GetFileNameWithoutExtension(templatePath) + "-" +
                    (from n in param where n.name == "MsgNumber" select n.value).First() + ".docx";

                try
                {
                    File.Copy(templatePath, newFilePath);

                    using (var outputDocument = new TemplateProcessor(newFilePath)
                        .SetRemoveContentControls(true))
                    {
                        outputDocument.FillContent(valuesToFill);
                        outputDocument.SaveChanges();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

        }
    }
}
