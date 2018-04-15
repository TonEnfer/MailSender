using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NameCaseLib;
using NameCaseLib.NCL;


namespace MailSender
{
    public static class Declator
    {
        private static Ru nc = new Ru();
        public static string Decline(string FullName, Padeg padeg)
        {
            nc.FullReset();
            String tmp = nc.Q(FullName, padeg);
            return tmp;
        }
    }
}
