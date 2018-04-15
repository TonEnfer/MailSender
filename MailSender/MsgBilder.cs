using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace MailSender
{
    public static class MsgFileBilder
    {
        struct Parameters
        {
            string name;
            string value;
        }

        static void Build(Parameters[] param, string patch) {

            throw new InvalidArgumentFormat();
        }
    }

    [Serializable()]
    public class InvalidArgumentFormat : System.ArgumentException
    {
        public InvalidArgumentFormat() : base() { }
        public InvalidArgumentFormat(string message) : base(message) { }
        public InvalidArgumentFormat(string message, System.Exception inner) : base(message, inner) { }

        // A constructor is needed for serialization when an
        // exception propagates from a remoting server to the client. 
        protected InvalidArgumentFormat(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
        { }
    }
}
