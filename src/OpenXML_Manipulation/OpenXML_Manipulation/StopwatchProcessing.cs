using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Manipulation
{
    class StopwatchProcessing
    {
        public static void Estimate()
        {
            Stopwatch sw = new Stopwatch();
            int no = 0;

            sw.Start();

            for (int i = 1; i < 100000; i++)
            {
                no += i;
            }

            sw.Stop();

            Console.WriteLine("Elapsed Seconds : " + sw.ElapsedMilliseconds / 1000);


        }
    }
}
