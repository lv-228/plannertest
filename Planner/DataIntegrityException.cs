using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planner
{
    class DataIntegrityException : Exception
    {
        public DataIntegrityException(string message)
        : base(message)
    { }
    }
}
