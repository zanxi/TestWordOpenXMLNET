using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Latex2MathML
{
    public class ExceptionEventArgs : EventArgs
    {
        public ExceptionEventArgs(String message)
        {
            this.Message = message;
        }

        public String Message { get; private set; }
    }
}
