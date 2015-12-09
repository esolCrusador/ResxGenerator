using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ResourcesAutogenerate
{
    public class ResxGenException:Exception
    {
        public ResxGenException(string message) : base(message)
        {
        }

        protected ResxGenException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public ResxGenException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public ResxGenException()
        {
        }
    }
}
