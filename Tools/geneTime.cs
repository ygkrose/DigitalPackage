using System;
using System.Collections.Generic;
using System.Text;

namespace DigitalSealed.Tools
{
    public static class geneTime
    {
        private static System.Diagnostics.Stopwatch sw1 = new System.Diagnostics.Stopwatch(); //偵測DCP執行速度
        static string counterString = "";
        //static string apath = @"c:\DigitWare\NPB-WEBOE\Debug\";

        public static string getTimeNow()
        {
            return DateTime.Now.AddYears(-1911).Year.ToString("000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00");               
        }

        //public static string getCounter(string mesg)
        //{
        //    if (mesg == "")
        //    {
        //        counterString = "";
        //        sw1.Reset();
        //        sw1.Start();
        //        return "";
        //    }
            
        //    if (mesg == "End")
        //    {
        //        if (System.IO.File.Exists(apath))
        //            System.IO.File.Delete(apath);
        //        using (System.IO.StreamWriter sw = new System.IO.StreamWriter(apath, false, Encoding.UTF8))
        //        {
        //            sw.Write(counterString);
        //        }
        //        counterString = "";
        //        apath = "";
        //    }
        //    counterString += mesg + ":" + ((float)sw1.ElapsedMilliseconds / 1000.0F).ToString() + "秒\r\n";
        //    sw1.Reset();
        //    sw1.Start();
        //    return counterString;
        //}

        //public static void getCounter(string mesg,string docno)
        //{
        //    apath += docno + "DCPLib計時.txt";
        //    getCounter(mesg);
        //}
    }
}
