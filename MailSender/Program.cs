using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace MailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            FileConverter fc = new FileConverter();
            MailVariable mv = new MailVariable();
            MailVariable.Reader mvr;


            try
            {
                Console.WriteLine("Чтение параметров из Excel...");
                mvr = new MailVariable.Reader("..//..//..//Рассылка//Список-1.xlsx");
                mv.AddVariables(mvr.ReadVariable());
                Console.WriteLine("Параметры прочитаны!");
                Console.WriteLine("Создание файлов...");
                foreach (var k in mv.GetVariables())
                {
                    List<MsgFileBilder.TextParameters> textParameters =
                        new List<MsgFileBilder.TextParameters>();
                    textParameters.Add(new MsgFileBilder.TextParameters("Organization", k.organization));
                    string io = k.fullName.Split(' ')[1].First() +
                        (string)". " + k.fullName.Split(' ')[2].First() + (string)". ";
                    string shName = io + (string)Declator.Decline(k.fullName, NameCaseLib.NCL.Padeg.DATELN).Split(' ')[0];
                    textParameters.Add(new MsgFileBilder.TextParameters("ShortName", shName));
                    textParameters.Add(new MsgFileBilder.TextParameters("CurrentDate", 
                        DateTime.Now.Date.ToShortDateString()));
                    textParameters.Add(new MsgFileBilder.TextParameters("MsgNumber",
                        k.number.ToString()));
                    textParameters.Add(new MsgFileBilder.TextParameters("Appeal",
                        Declator.getSex(k.fullName) == NameCaseLib.NCL.Gender.Man ? "Уважаемый" : "Уважаемая"));
                    textParameters.Add(new MsgFileBilder.TextParameters("LongName", k.fullName.Split(' ')[1] +
                        (string)" " + k.fullName.Split(' ')[2]));
                    MsgFileBilder.Build(textParameters, "..\\..\\..\\Рассылка\\Информационное письмо.docx");
                    
                }
                Console.WriteLine("Файлы созданы!");
                Console.WriteLine("Конвертация файлов...");
                foreach (var k in mv.GetVariables())
                {
                    string inputFile = "..\\..\\..\\Рассылка\\Информационное письмо" + "-" + k.number + ".docx";
                    string outputFile = "..\\..\\..\\Рассылка\\Информационное письмо" + "-" + k.number + ".pdf";
                    fc.Convert(inputFile, outputFile);
                }
                Console.WriteLine("Файлы сконвертированы...");
            }
            catch { }


            String a = Declator.Decline("Иванов Иван Иванович", NameCaseLib.NCL.Padeg.RODITLN);
            System.Console.WriteLine(a);
            Console.ReadKey();
        }
    }
}
