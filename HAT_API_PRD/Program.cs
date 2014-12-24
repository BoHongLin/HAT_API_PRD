using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAT_API_PRD
{
    class Program
    {
        static void Main(string[] args)
        {
            //你建立的flow
            PrdFlow flow = new PrdFlow();
            flow.DoPrdFlow();
        }
    }
}
