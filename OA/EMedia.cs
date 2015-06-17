using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature;
using DigitalSealed.Tools;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using FSXMLCAPIATLLib;
using FSGPKICRYPTATLLib;

namespace DigitalSealed.OA
{
    /// <summary>
    /// 媒體封裝
    /// </summary>
    public class EMedia
    {
        //詮釋資料檔
        private string _MataDatafullpath = "";

        private string _媒體封裝SignID = "ermedia";

        public string 媒體封裝SignID
        {
            get { return _媒體封裝SignID; }
            set { _媒體封裝SignID = value; }
        }

        private string _HashAlgorithm = "SHA1"; //or SHA256
        /// <summary>
        /// 選用雜湊演算法目前只支援SHA1, SHA256
        /// </summary>
        public string HashAlgorithm
        {
            get { return _HashAlgorithm; }
            set { _HashAlgorithm = value; }
        }

        //接收機關憑證檔位置
        private string _ReceivedOrgCert = "";

        private string _CipherValue = "";

        private X509Certificate2 signCert = null;
        private string _signCertSubjectName = "";

        private byte[] akey = null;
        private string _RawKey = "";
        string sKey = "";
        string sIV = "";


        public EMedia(string 金鑰信封檔, string 機關pfx檔, string pwd)
        {
            if (!File.Exists(金鑰信封檔)) { throw new Exception("金鑰信封檔不存在!!"); }
            XmlDocument xdoc = new XmlDocument();
            xdoc.XmlResolver = null;
            xdoc.Load(金鑰信封檔);
            if (xdoc.DocumentElement.SelectSingleNode("//CipherValue") != null)
                _CipherValue = xdoc.DocumentElement.GetElementsByTagName("CipherValue")[0].InnerText;
            if (_CipherValue == "") { throw new Exception("金鑰字串不存在!!"); }
            int FS_FLAG_HEX_ENCODE = 0x01000000;
            GPKICryptATLClass gpkixml = new GPKICryptATLClass();
            _RawKey = gpkixml.FSGPKI_RSADecrypt(_CipherValue, pwd, FS_FLAG_HEX_ENCODE);
            if (_RawKey == null) { throw new Exception("金鑰信封解密失敗!"); }
        }

        //public EMedia(string 金鑰信封檔,string 機關pfx檔,string pwd) 
        //{
        //    if (!File.Exists(金鑰信封檔)) { throw new Exception("金鑰信封檔不存在!!"); }
        //    XmlDocument xdoc = new XmlDocument();
        //    xdoc.XmlResolver = null;
        //    xdoc.Load(金鑰信封檔);
        //    if (機關pfx檔 != "")
        //    {
        //        signCert = new X509Certificate2(機關pfx檔, pwd, X509KeyStorageFlags.Exportable);
        //        //if (!(signCert.Verify())) { throw new Exception("憑證驗證失敗!!"); }
        //    }
        //    else
        //    {
        //        if (!loadcertfromCSP(xdoc.DocumentElement))
        //        {
        //            throw new Exception("找不到" + _signCertSubjectName + "的對應憑證可解密!!");
        //        }
        //    }
        //    if (xdoc.DocumentElement.SelectSingleNode("//CipherValue") != null)
        //        _CipherValue = xdoc.DocumentElement.GetElementsByTagName("CipherValue")[0].InnerText;
        //    if (_CipherValue == "") { throw new Exception("金鑰字串不存在!!"); }
            
        //    byte[] cv = Convert.FromBase64String(_CipherValue); 
        //    System.Security.Cryptography.RSACryptoServiceProvider drsa = signCert.PrivateKey as System.Security.Cryptography.RSACryptoServiceProvider;
        //    //System.Security.Cryptography.RSACryptoServiceProvider drsa = new RSACryptoServiceProvider(); 
            
        //    //drsa.FromXmlString(signCert.PrivateKey.ToXmlString(true));
        //    //byte[] akey = null;
        //    if (cv.Length == 128)
        //    {
        //        akey = drsa.Decrypt(cv, false);
        //    }
        //    else
        //    {
        //        Array.Reverse(cv);
        //        byte[] destinationArray = new byte[128];
        //        Array.Copy(cv, destinationArray, 128);
        //        akey = drsa.Decrypt(cv, false);
        //    }

        //    if (akey == null) { throw new Exception("金鑰解密失敗!!"); }
        //}

        public EMedia(string 詮釋資料檔) 
        {
            _MataDatafullpath = 詮釋資料檔;
            if (!File.Exists(_MataDatafullpath)) throw new Exception("詮釋檔不存在");
        }

        public EMedia(string 詮釋資料檔,string 接收機關憑證檔) 
        {
            _MataDatafullpath = 詮釋資料檔;
            _ReceivedOrgCert = 接收機關憑證檔;
            if (!File.Exists(_MataDatafullpath)) throw new Exception("詮釋檔不存在");
            if (!File.Exists(_ReceivedOrgCert)) throw new Exception("憑證檔不存在");
        }

        /// <summary>
        /// 產生媒體封裝檔並加簽
        /// </summary>
        /// <param name="pwd"></param>
        /// <returns></returns>
        public XmlDocument MakeEMediaXml(string pwd)
        {
            FileInfo metadata = new FileInfo(_MataDatafullpath);
            XmlDocument doc = new XmlDocument();
            doc.XmlResolver = null;
            doc.Load(_MataDatafullpath);

            List<string> refs = new List<string>();
            refs.Add(metadata.Name);

            foreach (XmlNode x in doc.DocumentElement.GetElementsByTagName("電子檔案名稱"))
            {
                FileInfo f = new FileInfo(metadata.Directory.FullName + "\\" + Path.Combine(x.PreviousSibling.InnerText, x.InnerText));
                if (f.Exists)
                {
                    //refs.Add(file[0].DirectoryName.Replace(metadata.DirectoryName + "\\", "") + "\\" + file[0].Name);
                    refs.Add(Path.Combine(f.DirectoryName.Replace(metadata.DirectoryName, ""), f.Name));
                }
                else
                    throw new Exception("檔案:" + x.InnerText + "路徑有誤!");
            }

            SignatureTag st = new SignatureTag(_媒體封裝SignID);
            st.HashAlgorithm = _HashAlgorithm;
            XmlNode sig = st.getSignatureNode(refs);

            XmlDocument DcpXml = new XmlDocument();
            DcpXml.XmlResolver = null;
            DcpXml.PreserveWhitespace = true;
            DcpXml.AppendChild(DcpXml.CreateXmlDeclaration("1.0", "utf-8", null));
            DcpXml.AppendChild(DcpXml.CreateDocumentType("媒體封裝檔", null, "99_ermedia_utf8.dtd", null));
            XmlNode media = xmlTool.MakeNode("媒體封裝檔", "");

            media.AppendChild(sig);
            DcpXml.AppendChild(DcpXml.ImportNode(media, true));
            DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature sign = new SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature(DcpXml);
            sign.FilePath = metadata.DirectoryName;
            if (_ReceivedOrgCert != "")
            {
                AddRcvOrgKey(ref DcpXml);
            }
            return sign.getSignXmlData(_媒體封裝SignID, pwd);
        }


        /// <summary>
        /// 置入金鑰信封並加密檔案
        /// </summary>
        /// <param name="rcvnd"></param>
        /// <returns></returns>
        private XmlDocument AddRcvOrgKey(ref XmlDocument rcvnd)
        {
            X509Certificate2 rcvcert = new X509Certificate2(_ReceivedOrgCert);
            XmlNode _受移轉機關金鑰信封 = xmlTool.MakeNode("受移轉機關金鑰信封", "");
            TripleDESCryptoServiceProvider des = new TripleDESCryptoServiceProvider();
            _受移轉機關金鑰信封.AppendChild(_受移轉機關金鑰信封.OwnerDocument.ImportNode(Envelope(rcvcert, ref des),true));
            encryptFile(des);
            rcvnd.DocumentElement.InsertBefore(rcvnd.ImportNode(_受移轉機關金鑰信封,true), rcvnd.DocumentElement.GetElementsByTagName("Signature")[0]);
            return rcvnd;
        }
        
        /// <summary>
        /// 產生金鑰信封內容
        /// </summary>
        /// <param name="RecvCert">接收方憑證</param>
        /// <returns></returns>
        private XmlElement Envelope(X509Certificate2 RecvCert, ref TripleDESCryptoServiceProvider _des)
        {
            try
            {
                CipherMode _mode = CipherMode.CBC;
                PaddingMode _pad = PaddingMode.PKCS7;
                //Encoding
                Encoding ecode = new System.Text.UnicodeEncoding();
                //Encoding ecode = new System.Text.UTF8Encoding();

                // Create the RSA CSP passing CspParameters object
                
                //rsa.ImportParameters(rp.ExportParameters(false));
                //Create a new AES Session Key 
                //RijndaelManaged _aes = new RijndaelManaged();
                _des = new TripleDESCryptoServiceProvider();
                _des.Padding = _pad;
                _des.Mode = _mode;

                //we need to send both the secret key(128 bits) and the IV(128 bits)
                //string sKey = ecode.GetString(_des.Key, 0, _des.Key.Length); //encoded to 16 character unicode string
                //string sIV = ecode.GetString(_des.IV, 0, _des.IV.Length);//encoded to 8 character unicode string
                sKey = ByteArrayToString(_des.Key);
                sIV = ByteArrayToString(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
                //Concatinate the secret key and IV to build a session key
                //StringBuilder sesskeyBuilder = new StringBuilder(sKey);
                //sesskeyBuilder.Append(sIV);
                //string sessionkey = sesskeyBuilder.ToString(); //24 chacacter unicode string

                //Envelope the Session key with Bob's Public Keys
                RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
                rsa.FromXmlString(RecvCert.PublicKey.Key.ToXmlString(false));
                byte[] envelope = rsa.Encrypt(StringToByteArray(sKey), false);

                EncryptedKey ek = new EncryptedKey();
                EncryptionMethod emthd = new EncryptionMethod("http://www.w3.org/2001/04/xmlenc#rsa-1_5");
                KeyInfo kinfo = new KeyInfo();
                kinfo.AddClause(new KeyInfoX509Data(RecvCert));
                CipherData cdt = new CipherData(envelope);
                ek.CipherData = cdt;
                ek.EncryptionMethod = emthd;
                ek.KeyInfo = kinfo;
                return ClearXmlns(ek.GetXml());
            }
            catch (Exception err)
            {
                throw new Exception("產生對稱式金鑰失敗!" + err.Message);
            }
        }

        public static string ByteArrayToString(byte[] ba)
        {
            string hex = BitConverter.ToString(ba);
            return hex.Replace("-", "");
        }

        public static byte[] StringToByteArray(String hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }

        //拿掉xmlns屬性
        private XmlElement ClearXmlns(XmlElement inxlm)
        {
            XmlDocument reDoc = new XmlDocument();
            inxlm.InnerXml = inxlm.InnerXml.Replace("xmlns=\"" + inxlm.NamespaceURI + "\"", "");
            inxlm.InnerXml = inxlm.InnerXml.Replace(" xmlns=\"http://www.w3.org/2000/09/xmldsig#\"", "");
            //EncryptionMethod emthd = new EncryptionMethod("http://www.w3.org/2001/04/xmlenc#tripledes-cbc");
            //string Method=emthd.GetXml().OuterXml.Replace("xmlns=\"" + inxlm.NamespaceURI + "\" ", ""); 
            reDoc.LoadXml("<" + inxlm.Name + ">" + inxlm.InnerXml + "</" + inxlm.Name + ">");
            return reDoc.FirstChild as XmlElement;
        }

        private bool loadcertfromCSP(XmlElement envele)
        {
            X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String((envele.GetElementsByTagName("X509Certificate")[0] as XmlElement).InnerText));
            _signCertSubjectName = certificate.Subject;
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.MaxAllowed);
            X509Certificate2Enumerator enumerator = store.Certificates.GetEnumerator();
            while (enumerator.MoveNext())
            {
                X509Certificate2 current = enumerator.Current;
                if (current.Thumbprint.Equals(certificate.Thumbprint))
                {
                    signCert = current;
                    //certificate2 = current;
                    break;
                }
            }
            if (signCert == null)
                return false;
            else
                return true;
        }

        /// <summary>
        /// 對檔案加密
        /// </summary>
        /// <param name="_des">對稱式金鑰</param>
        private void encryptFile(TripleDESCryptoServiceProvider _des)
        {
            try
            {
                DirectoryInfo dir = new FileInfo(_MataDatafullpath).Directory;
                foreach (FileInfo f in dir.GetFiles("*.*", SearchOption.AllDirectories))
                {
                    if (f.Extension.ToLower() == ".xml") continue;
                    FileStream outStream = File.OpenWrite(Path.Combine(f.DirectoryName,"encrytmp"));
                    ICryptoTransform transform = _des.CreateEncryptor(StringToByteArray(sKey),StringToByteArray(sIV));
                    CryptoStream cryptoStream = new CryptoStream(outStream, transform, CryptoStreamMode.Write);
                    string targetname = f.FullName;
                    Byte[] inFile = File.ReadAllBytes(targetname);
                    cryptoStream.Write(inFile, 0, inFile.Length);
                    cryptoStream.Close();
                    outStream.Close();
                    f.Delete();
                    File.Move(Path.Combine(f.DirectoryName, "encrytmp"), targetname); 
                }
            }
            catch (Exception err)
            {
                throw new Exception("檔案加密失敗:" + err.Message);
            }
        }

        /// <summary>
        /// 檔案解密
        /// </summary>
        /// <param name="inPath"></param>
        /// <param name="outPath"></param>
        public void DecryptFile(String inPath, String outPath)
        {
            FSXMLCAPIATLLib.FSXMLATLClass xmlcapi = new FSXMLATLClass();
            int FS_ALGOR_3DES = 0x02;
            string iv = "0000000000000000";
            int decRtn = xmlcapi.FSXMLCAPI_SymmetricRawKeyDecryptFile2File(FS_ALGOR_3DES, _RawKey, iv , inPath, outPath);
            if (decRtn != 0)
            {
                throw new Exception(ErrorMsgTable.GetErrMessage(decRtn)); 
            }
            else
                return ;
            #region use CSP
            //CspParameters cspParams = new CspParameters(1);
            //cspParams.KeyContainerName = "mycontainername";
            //cspParams.KeyNumber = 1;
            //cspParams.ProviderName = "Microsoft Base Cryptographic Provider v1.0";
            //cspParams.Flags = CspProviderFlags.UseMachineKeyStore;
            //RSACryptoServiceProvider clientRSA=null;
            //try
            //{
            //clientRSA = new RSACryptoServiceProvider( cspParams );
            //clientRSA.PersistKeyInCsp = true;
            //}
            //catch( Exception ex )
            //{
            //System.Diagnostics.Debug.WriteLine( ex.Message );

            //}

            //RSAPKCS1KeyExchangeDeformatter pRSADef = new
            //RSAPKCS1KeyExchangeDeformatter( clientRSA );
            //RijndaelManaged rijndael = new RijndaelManaged();
            //try
            //{
            //    rijndael.Key = pRSADef.DecryptKeyExchange(ChiperValue);
            //}
            //catch (Exception ex)
            //{
            //}
            #endregion


            if (akey == null) { throw new Exception("無金鑰可解密!!"); }

            Stream inStream = null;
            CryptoStream cryptoStream;
            if (akey.Length == 32)
            {
                Encoding ecode = new UnicodeEncoding();
                string keystr = ecode.GetString(akey);
                string sKey = keystr.Substring(0, 12);
                string sIV = keystr.Substring(12, 4);

                SymmetricAlgorithm symmetricAlgorithm = SymmetricAlgorithm.Create("TripleDES");
                symmetricAlgorithm.Mode = CipherMode.CBC;
                symmetricAlgorithm.Padding = PaddingMode.PKCS7;
                symmetricAlgorithm.Key = ecode.GetBytes(sKey);
                symmetricAlgorithm.IV = ecode.GetBytes(sIV);
                inStream = File.OpenRead(inPath);
                ICryptoTransform transform = symmetricAlgorithm.CreateDecryptor();
                cryptoStream = new CryptoStream(inStream, transform, CryptoStreamMode.Read);
            }
            else if (akey.Length == 24)
            {
                TripleDESCryptoServiceProvider provider = new TripleDESCryptoServiceProvider();
                byte[] rgbIV = new byte[8];
                inStream = File.OpenRead(inPath);
                cryptoStream = new CryptoStream(inStream, provider.CreateDecryptor(akey, rgbIV), CryptoStreamMode.Read);
            }
            else
            {
                throw new Exception("密鑰長度錯誤!!");
            }

            
            //Byte[] buffer = new Byte[1];
            Byte[] buffer = new Byte[cryptoStream.Length];
            int length = cryptoStream.Read(buffer, 0, buffer.Length);

            Stream outStream = File.OpenWrite(outPath);

            //StreamWriter fsPlaintext = new StreamWriter(outPath);
            //fsPlaintext.Write(new StreamReader(cryptoStream).ReadToEnd());
            //fsPlaintext.Flush();
            //fsPlaintext.Close(); 
            while (length > 0)
            {
                outStream.Write(buffer, 0, length);
                length = cryptoStream.Read(buffer, 0, buffer.Length);
            }

            inStream.Close();
            outStream.Close();
        }

     
    }
}
