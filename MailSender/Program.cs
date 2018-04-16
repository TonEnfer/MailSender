using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
                    string io = k.fullName.Split(' ')[1].First() + ". " +
                        k.fullName.Split(' ')[2].First() + ". ";
                    string shName = io + (string)Declator.Decline(k.fullName, NameCaseLib.NCL.Padeg.DATELN).Split(' ')[0];
                    List<MsgFileBilder.TextParameters> textParameters =
                        new List<MsgFileBilder.TextParameters>
                        {
                            new MsgFileBilder.TextParameters("Organization", k.organization),
                            new MsgFileBilder.TextParameters("ShortName", shName),
                            new MsgFileBilder.TextParameters("CurrentDate",
                        DateTime.Now.Date.ToShortDateString()),
                            new MsgFileBilder.TextParameters("MsgNumber",
                        k.number.ToString()),
                            new MsgFileBilder.TextParameters("Appeal",
                        Declator.getSex(k.fullName) == NameCaseLib.NCL.Gender.Man ? "Уважаемый" : "Уважаемая"),
                            new MsgFileBilder.TextParameters("LongName", k.fullName.Split(' ')[1] +
                        (string)" " + k.fullName.Split(' ')[2])
                        };

                    MsgFileBilder.Build(textParameters, "..\\..\\..\\Рассылка\\Информационное письмо.docx");
                    MsgFileBilder.Build(textParameters, "..\\..\\..\\Рассылка\\Текст письма.txt");

                    GC.Collect();
                }
                Console.WriteLine("Файлы созданы!");
                Console.WriteLine("Конвертация файлов...");
                foreach (var k in mv.GetVariables())
                {
                    try
                    {
                        string inputFile = "..\\..\\..\\Рассылка\\Generated\\Информационное письмо" + "-" + k.number + ".docx";
                        string outputFile = "..\\..\\..\\Рассылка\\Generated\\Информационное письмо" + "-" + k.number + ".pdf";
                        if (!File.Exists(outputFile))
                            fc.Convert(inputFile, outputFile);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
                Console.WriteLine("Файлы сконвертированы!");
                Console.WriteLine("Отправка писем...");
                foreach (var k in mv.GetVariables())
                {
                    try
                    {
                        Sender.SendMail("server",
                                "login",
                                "Name",
                                "password",
                                k.email,
                                "Целевой приём-2018",
                                File.ReadAllText("..\\..\\..\\Рассылка\\Generated\\Текст письма-" + k.number + ".txt"),
                                new List<string> {
                                "..\\..\\..\\Рассылка\\Generated\\Информационное письмо" + "-" + k.number + ".pdf",
                                "..\\..\\..\\Рассылка\\Заявка о целевом приёме.docx" }
                                );
                        Thread.Sleep(100);
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine("Письмо не отправлено : {0}", e.Message);
                    }
                    
                }
            }
            catch { }
            Console.WriteLine("Для завершения нажмите любую кнопку");
            Console.ReadKey();
        }
    }
}
