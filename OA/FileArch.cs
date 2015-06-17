using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using DigitalSealed.Tools;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature;

namespace DigitalSealed.OA
{
    /// <summary>
    /// 檔管異動封裝檔
    /// </summary>
    public class Archivist
    {
        private string _電子封裝檔位置 = "";
        private string _封裝檔檔名 = "";
        private XmlDocument sourceXml = new XmlDocument();
        private List<string> refNodes = new List<string>();

        private string _詮釋資料SignID = "";

        public string 詮釋資料SignID
        {
            get { return _詮釋資料SignID; }
            set { _詮釋資料SignID = value; }
        }

        private string _PKCS11Driver = "";

        public string PKCS11Driver
        {
            get { return _PKCS11Driver; }
            set { _PKCS11Driver = value; }
        }

        private string _HashAlgorithm = "SHA1"; //or SHA256
        /// <summary>
        /// 選用雜湊演算法目前只支援SHA1, SHA256
        /// </summary>
        public string HashAlgorithm
        {
            get { return _HashAlgorithm; }
            set { _HashAlgorithm = value;}
        }

        public Archivist() { }

        /// <summary>
        /// 檔管點收建構子
        /// </summary>
        /// <param name="filelocation">簽核電子檔(si檔)</param>
        /// <param name="outputfilefullname">輸出檔案檔名及路徑</param>
        public Archivist(string filelocation,string outputfilefullname)
        {
            if (File.Exists(filelocation))
            {
                sourceXml.XmlResolver = null;
                sourceXml.Load(filelocation);
                sourceXml.PreserveWhitespace = true;
            }
            else
                throw new Exception("簽核電子檔不存在!");
            _電子封裝檔位置 = outputfilefullname;
        }

        /// <summary>
        /// 封裝詮釋資料建構子
        /// </summary>
        /// <param name="filelocation">電子封裝檔路徑</param>
        public Archivist(string filelocation)
        {
            
            FileInfo fi = new FileInfo(filelocation); 
            if (fi.Exists)
            {
                _電子封裝檔位置 = fi.FullName;
                sourceXml.XmlResolver = null;
                sourceXml.Load(filelocation);
                sourceXml.PreserveWhitespace = true;
            }
            else
                throw new Exception("電子封裝檔不存在!");
        }

        /// <summary>
        /// 影音封裝專用建構子
        /// </summary>
        /// <param name="filelocation">影音封裝檔路徑</param>
        /// <param name="refnds">要算Hash的影像檔</param>
        public Archivist(string filelocation,List<string> refnds)
        {
            FileInfo fi = new FileInfo(filelocation);
            if (fi.Exists)
            {
                _電子封裝檔位置 = fi.FullName;
                sourceXml.XmlResolver = null;
                sourceXml.Load(filelocation);
                sourceXml.PreserveWhitespace = true;
                refNodes = refnds;
            }
            else
                throw new Exception("電子封裝檔不存在!");
        }

        /// <summary>
        /// 檔管點收加機關憑證
        /// </summary>
        /// <param name="pwd">機關憑證密碼</param>
        public string 檔管點收加機關憑證(string pwd)
        {
            try
            {
                //先產生電子封裝檔的空殼
                newElecDcpXmlFile();
                FileInfo tmp = new FileInfo(_電子封裝檔位置);
                Signature sign = new Signature(tmp.DirectoryName, tmp.Name);
                if (_PKCS11Driver != "") sign.PKCS11Driver = _PKCS11Driver;
                XmlDocument doc = sign.getSignXmlData("CheckSignGCA", pwd);
                doc.Save(_電子封裝檔位置);
                return "";
            }
            catch (Exception err)
            {
                return err.Message; 
            }
        }

        public string 前端執行點收(string siString,string pwd)
        {
            try
            {
                XmlDocument sidoc = new XmlDocument();
                sidoc.XmlResolver = null;
                sidoc.LoadXml(siString);
                sidoc.PreserveWhitespace = true;
                Signature sign = new Signature(sidoc);
                if (_PKCS11Driver != "") sign.PKCS11Driver = _PKCS11Driver;
                XmlDocument doc = sign.SignOnlyXmlData("CheckSignGCA", pwd);
                doc.Save(_電子封裝檔位置);
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        public string 取得檔管點收加簽前資料()
        {
            try
            {
                //先產生電子封裝檔的空殼
                newElecDcpXmlFile();
                FileInfo tmp = new FileInfo(_電子封裝檔位置);
                Signature sign = new Signature(tmp.DirectoryName, tmp.Name);
                if (_PKCS11Driver != "") sign.PKCS11Driver = _PKCS11Driver;
                return sign.getHashXmlData("CheckSignGCA");
            }
            catch (Exception err)
            {
                return "err:" + err.Message;
            }
        }

        /// <summary>
        /// 加入詮釋資料
        /// </summary>
        /// <param name="pwd">詮釋資料</param>
        /// <param name="pwd">機關憑證密碼,空白代表只算hash不加簽</param>
        public string 加入詮釋資料(XmlNode descdata,string pwd)
        {
            try
            {
                //判斷封裝檔電子簽章存在時先移除
                XmlNodeList tagnd = sourceXml.DocumentElement.GetElementsByTagName("封裝檔電子簽章");
                while (tagnd.Count > 0)
                {
                    sourceXml.DocumentElement.RemoveChild(sourceXml.DocumentElement.GetElementsByTagName("封裝檔電子簽章")[0]); 
                }
                //加入<封裝檔電子簽章>
                if (sourceXml.DocumentElement.GetElementsByTagName("電子影音檔案").Count > 0)
                {
                    if (descdata != null)
                    {
                        refNodes.Clear();
                        foreach (XmlNode xnd in sourceXml.DocumentElement.SelectNodes("//電子影音檔案資訊/檔案名稱"))
                            refNodes.Add(xnd.InnerText);
                        //影音封裝
                        sourceXml.DocumentElement.InsertBefore(sourceXml.ImportNode(genepackageElecSign(refNodes), true), sourceXml.DocumentElement.GetElementsByTagName("封裝檔內容")[0]);
                    }
                }
                else
                {
                    //電子檔案封裝
                    sourceXml.DocumentElement.InsertBefore(sourceXml.ImportNode(genepackageElecSign(), true), sourceXml.DocumentElement.GetElementsByTagName("封裝檔內容")[0]);
                }
                //判斷詮釋資料存在時先移除
                XmlNodeList nds = sourceXml.DocumentElement.SelectNodes("封裝檔內容/詮釋資料");
                if (nds.Count > 0)
                {
                    foreach (XmlNode xnd in nds)
                        xnd.ParentNode.RemoveChild(xnd); 
                }
                //加入詮釋資料
                if (descdata != null)
                {
                    sourceXml.DocumentElement.LastChild.InsertAfter(sourceXml.ImportNode(descdata, true), sourceXml.DocumentElement.LastChild.FirstChild);
                    sourceXml.Save(_電子封裝檔位置);
                    FileInfo tmp = new FileInfo(_電子封裝檔位置);
                    Signature sign = new Signature(tmp.DirectoryName, tmp.Name);
                    if (_PKCS11Driver != "") sign.PKCS11Driver = _PKCS11Driver;
                    XmlNode snd = sourceXml.DocumentElement.SelectSingleNode("//封裝檔電子簽章/Signature");
                    if (snd.Attributes["Id"] == null) return "找不到 //封裝檔電子簽章/Signature!";
                    _詮釋資料SignID = snd.Attributes["Id"].Value;
                    if (pwd == "")
                    {
                        sourceXml.LoadXml(sign.getHashXmlData(snd.Attributes["Id"].Value));
                        sourceXml.Save(_電子封裝檔位置);
                    }
                    else
                    {
                        sign.getSignXmlData(snd.Attributes["Id"].Value, pwd).Save(_電子封裝檔位置);
                    }
                }
                else
                    sourceXml.Save(_電子封裝檔位置);
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        /// <summary>
        /// 產生新的電子封裝檔
        /// </summary>
        private void newElecDcpXmlFile()
        {
            XmlDocument DcpXml = new XmlDocument();
            DcpXml.XmlResolver = null;
            DcpXml.PreserveWhitespace = true;
            DcpXml.AppendChild(DcpXml.CreateXmlDeclaration("1.0", "utf-8", null));
            DcpXml.AppendChild(DcpXml.CreateDocumentType("電子封裝檔", null, "99_erencaps_utf8.dtd", null));
             
            XmlNode 封裝檔內容 = xmlTool.MakeNode("封裝檔內容", "Id", "Wrap");
            封裝檔內容.AppendChild(xmlTool.MakeNode("封裝檔資訊", "電子檔案"));
            封裝檔內容.AppendChild(xmlTool.MakeNode("電子檔案", ""));
            XmlNode 電子封裝檔 = xmlTool.MakeNode("電子封裝檔", "");
            電子封裝檔.AppendChild(封裝檔內容);
            DcpXml.AppendChild(DcpXml.ImportNode(電子封裝檔, true));
            //匯入檔案管理單位點收簽章節點
            DcpXml.DocumentElement.GetElementsByTagName("電子檔案")[0].AppendChild(DcpXml.ImportNode(geneRcvSignNode(), true));
            //匯入線上簽核節點內容
            DcpXml.DocumentElement.GetElementsByTagName("電子檔案")[0].AppendChild(DcpXml.ImportNode(sourceXml.GetElementsByTagName("線上簽核")[0], true));
            
            DcpXml.Save(_電子封裝檔位置); 
        }

        /// <summary>
        /// 產生空的檔案管理單位點收簽章
        /// </summary>
        /// <returns></returns>
        private XmlNode geneRcvSignNode()
        {
            XmlNode rtn = xmlTool.MakeNode("檔案管理單位點收簽章", "");
            SignatureTag st = new SignatureTag("CheckSignGCA");
            st.HashAlgorithm = _HashAlgorithm;
            List<string> refnd = new List<string>();
            refnd.Add("#FlowInfo");
            string fid = sourceXml.DocumentElement.GetElementsByTagName("線上簽核資訊")[0].Attributes["Id"].Value ; 
            if (fid != "")
                refnd.Add("#" + fid);
            else
                refnd.Add("#SignInfo");
            refnd.Add("#CheckSignByGcaTime");
            rtn.AppendChild(st.getSignatureNode(refnd));
            XmlNode 加簽時間 = xmlTool.MakeNode("加簽時間",geneTime.getTimeNow());
            XmlAttribute xatt = 加簽時間.OwnerDocument.CreateAttribute("Id");
            xatt.Value = "CheckSignByGcaTime";
            加簽時間.Attributes.Append(xatt);
            rtn.AppendChild(加簽時間);
            return rtn;
        }

        /// <summary>
        /// 產生封裝檔電子簽章
        /// </summary>
        /// <returns></returns>
        private XmlNode genepackageElecSign()
        {
            XmlNode rtn = xmlTool.MakeNode("封裝檔電子簽章", "");
            SignatureTag st = new SignatureTag("EnveSign");
            st.HashAlgorithm = _HashAlgorithm;
            List<string> refnd = new List<string>();
            refnd.Add("#Wrap");
            refnd.Add("#CheckEnveSignTime");
            rtn.AppendChild(st.getSignatureNode(refnd));
            XmlNode 加簽時間 = xmlTool.MakeNode("加簽時間", geneTime.getTimeNow());
            XmlAttribute xatt = 加簽時間.OwnerDocument.CreateAttribute("Id");
            xatt.Value = "CheckEnveSignTime";
            加簽時間.Attributes.Append(xatt);
            rtn.AppendChild(加簽時間);
            return rtn;
        }

        private XmlNode genepackageElecSign(List<string> refnd)
        {
            XmlNode rtn = xmlTool.MakeNode("封裝檔電子簽章", "");
            SignatureTag st = new SignatureTag("EnvelopeDS");
            st.HashAlgorithm = _HashAlgorithm;
            //List<string> refnd = new List<string>();
            refnd.Add("#Wrap");
            refnd.Add("#CheckEnveSignTime");
            rtn.AppendChild(st.getSignatureNode(refnd));
            XmlNode 加簽時間 = xmlTool.MakeNode("加簽時間", geneTime.getTimeNow());
            XmlAttribute xatt = 加簽時間.OwnerDocument.CreateAttribute("Id");
            xatt.Value = "CheckEnveSignTime";
            加簽時間.Attributes.Append(xatt);
            rtn.AppendChild(加簽時間);
            return rtn;
        }


    }
}
