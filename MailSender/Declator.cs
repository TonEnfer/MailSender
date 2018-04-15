using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
<<<<<<< HEAD
using NameCaseLib;
=======
>>>>>>> 014b3f40376cb5b4f9bb5c1939b45d7c0a7dda51
using NameCaseLib.NCL;


namespace MailSender
{
    public static class Declator
    {
<<<<<<< HEAD
        private static Ru nc = new Ru();
        public static string Decline(string FullName, Padeg padeg)
        {
            nc.FullReset();
            String tmp = nc.Q(FullName, padeg);
            return tmp;
=======
        public static string Decline(string FullName, Padeg padeg)
        {
            String[] tmp = FullName.Split(' ');

            return "";
>>>>>>> 014b3f40376cb5b4f9bb5c1939b45d7c0a7dda51
        }
    }
}
