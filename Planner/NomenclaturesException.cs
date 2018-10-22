using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planner
{
    class NomenclaturesException : Exception
    {
        public NomenclaturesException(string message)
        : base(message)
    { }
    }
}
