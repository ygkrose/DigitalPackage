using System;
using System.Collections.Generic;
using System.Text;

namespace DigitalSealed
{
    /// <summary>
    /// 存放程式的共用變數
    /// </summary>
    public static class ConstVar
    {
        private static string _fileWorkingPath = @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\testDigitalSealed\bin\Debug\";

        public static string FileWorkingPath
        {
            get { return _fileWorkingPath; }
            set { _fileWorkingPath = value; }
        }

        private static string _sourceFilesPath = "";

        public static string SourceFilesPath
        {
            get { return ConstVar._sourceFilesPath; }
            set { ConstVar._sourceFilesPath = value; }
        }

        private static string _DTDPath = "";

        public static string DTDPath
        {
            get { return _DTDPath; }
            set { _DTDPath = value; }
        }
    }
}
