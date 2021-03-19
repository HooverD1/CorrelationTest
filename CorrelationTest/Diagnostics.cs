using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;

namespace CorrelationTest
{
    public static class Diagnostics
    {
        private static Stopwatch MyStopwatch { get; set; } = new Stopwatch();

        public static void StartTimer()
        {
            MyStopwatch.Reset();
            MyStopwatch.Start();
        }

        public static long CheckTimer()
        {
            return MyStopwatch.ElapsedMilliseconds;
        }

        public static void StopTimer(string message="", bool showTime = false)
        {
            MyStopwatch.Stop();
            if (showTime)
                MessageBox.Show($"{message}: {MyStopwatch.ElapsedMilliseconds.ToString()}");
        }
    }
}
