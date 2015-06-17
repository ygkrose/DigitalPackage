using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace DigitalSealed.Tools
{
    public static class xmlTool 
    {
        private static XmlDocument xDom = new XmlDocument(); 
        public static XmlNode MakeNode(string nodename,Dictionary<string,string> attb_value)
        {
            xDom.PreserveWhitespace = true;
            XmlElement xe1 = xDom.CreateElement(nodename);
            foreach (string akey in attb_value.Keys)
            {
                string avalue = "";
                attb_value.TryGetValue(akey, out avalue);
                XmlAttribute attb = xDom.CreateAttribute(akey);
                attb.Value = avalue;
                xe1.Attributes.Append(attb); 
            }
            return xe1 as XmlNode;
        }

        public static XmlNode MakeNode(string nodename, string attb, string attbvalue)
        {
            Dictionary<string, string> _attb = new Dictionary<string, string>();
            _attb.Add(attb, attbvalue);
            return MakeNode(nodename, _attb);  
        }

        public static XmlNode MakeNode(string nodename, string nodeValue)
        {
            xDom.PreserveWhitespace = true;
            XmlElement xe1 = xDom.CreateElement(nodename);
            xe1.InnerText = nodeValue;
            return xe1;
        }

    }
}
