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
            //fc.Convert("..//..//..//Рассылка//Информационное письмо.docx",
            //    "..//..//..//Рассылка//Информационное письмо.pdf");
            //fc.Free();
            MailVariable mv = new MailVariable();
            
            try
            {
                MailVariable.Reader mvr = new MailVariable.Reader("..//..//..//Рассылка//Список-1.xlsx");
                mv.AddVariables(mvr.ReadVariable());

                foreach (var k in mv.GetVariables())
                {
                    Console.Write("{0}\t{1}\t{2}\t", k.number, k.organization, k.fullName);
                    foreach (var s in k.email.ToArray())
                    {
                        Console.Write("{0}\t", s);
                    }
                    Console.WriteLine();
                }
            }
            catch { }


            String a = Declator.Decline("Иванов Иван Иванович", NameCaseLib.NCL.Padeg.RODITLN);
            System.Console.WriteLine(a);
            Console.ReadKey();
        }
    }
}
