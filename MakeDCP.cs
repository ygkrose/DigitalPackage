using System;
using System.Collections.Generic;
using System.Text;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo;
using System.Xml;
using System.IO;
using DigitalSealed.Tools;

namespace DigitalSealed
{
    /// <summary>
    /// 封裝檔主類別
    /// </summary>
    public class MakeDCP 
    {
        //線上簽核類別
        private SignOnline.SignOnline SO = null;

        /// <summary>
        /// 取得目前簽核點ID
        /// </summary>
        public string 簽核點ID
        {
            get
            {
                if (SO != null)
                    return SO.SignID;
                else
                    return "";
            }
        }

        private string _文號 = "";

        public string 公文文號
        {
            get 
            {
                if (_文號 != "")
                    return _文號;
                else if (SO != null)
                    return SO.DocNO;
                else
                    return "";
            }
        }

        private string _檔案來源資料夾 = "";
        /// <summary>
        /// 存放相關di及附件的資料夾
        /// </summary>
        public string 檔案來源資料夾
        {
            get { return _檔案來源資料夾; }
            set
            {
                _檔案來源資料夾 = value;
                ConstVar.SourceFilesPath = _檔案來源資料夾; 
            }
        }

        public XmlDocument 封裝檔XML
        {
            get { return SO.線上簽核; }
        }

        private string _暫存資料夾 = @"c:\temp\";

        public string 暫存資料夾
        {
            get { return _暫存資料夾; }
            set
            {
                _暫存資料夾 = value;
                ConstVar.FileWorkingPath = _暫存資料夾;
            }
        }

        private string _PKCS11Driver = "";
        /// <summary>
        /// 宣告引用其他憑證類別
        /// </summary>
        public string PKCS11Driver
        {
            set { _PKCS11Driver = value; }
        }

        private string _HashAlgorithm = "SHA1"; //or SHA256
        /// <summary>
        /// 選用雜湊演算法目前只支援SHA1, SHA256
        /// </summary>
        public string HashAlgorithm
        {
            get { return _HashAlgorithm; }
            set { _HashAlgorithm = value; SO.HashAlgorithm = _HashAlgorithm; }
        }

        private string _CertString = "";
        /// <summary>
        /// 由外部傳進憑證字串
        /// </summary>
        public string CertString
        {
            get { return _CertString; }
            set { _CertString = value; }
        }

        private bool _UseTifAsPage = true;
        /// <summary>
        /// 使用單一tif檔當頁面描述檔
        /// </summary>
        public bool UseTifAsPage
        {
            get { return _UseTifAsPage; }
            set { _UseTifAsPage = value; }
        }


        /// <summary>
        /// 建構子
        /// </summary>
        public MakeDCP(string DocNO)
        {
            SO = new SignOnline.SignOnline();
            SO.DocNO = DocNO;
        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="文號"></param>
        /// <param name="簽核人員"></param>
        /// <param name="次位簽核人員"></param>
        /// <param name="簽核人員意見"></param>
        /// <param name="異動類別"></param>
        public MakeDCP(string 文號, ModifyInfo.signer 簽核人員, ModifyInfo.signer 次位簽核人員, string 簽核人員意見, FlowInfo.FlowType 異動類別) 
        {
            _文號 = 文號;
            SO = new SignOnline.SignOnline(文號, 簽核人員, 異動類別);
            if (次位簽核人員 != null)
                SO.Receiver = 次位簽核人員;
            SO.SignerComment = 簽核人員意見;
        }

        /// <summary>
        /// 建構子(補簽專用)
        /// </summary>
        /// <param name="siXMLDoc">傳入已存在的封裝檔</param>
        /// <param name="簽核人員"></param>
        public MakeDCP(XmlDocument siXMLDoc, ModifyInfo.signer 簽核人員)
        {
            SO = new SignOnline.SignOnline(siXMLDoc, 簽核人員, FlowInfo.FlowType.補簽);
        }

        public MakeDCP(FileInfo 封裝檔位置,DirectoryInfo 解壓路徑, ModifyInfo.signer 簽核人員, ModifyInfo.signer 次位簽核人員, string 簽核人員意見, FlowInfo.FlowType 異動類別)
        {
            if (封裝檔位置.Extension.ToLower() == ".si")
            {
                XmlDocument xdoc = new XmlDocument();
                xdoc.XmlResolver = null;
                xdoc.Load(封裝檔位置.FullName);
                xdoc.PreserveWhitespace = true;
                this.檔案來源資料夾 = 解壓路徑.FullName;
                SO = new SignOnline.SignOnline(xdoc, 簽核人員, 異動類別);
                if (次位簽核人員 != null)
                    SO.Receiver = 次位簽核人員;
                SO.SignerComment = 簽核人員意見;
            }
            else if (this.解壓封裝檔(封裝檔位置.FullName, 解壓路徑.FullName))
            {
                DirectoryInfo dinfo = new DirectoryInfo(解壓路徑.FullName);
                if (dinfo.GetFiles("*.si").Length > 0)
                {
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.XmlResolver = null;
                    xdoc.Load(dinfo.GetFiles("*.si")[0].FullName);
                    xdoc.PreserveWhitespace = true;
                    this.檔案來源資料夾 = 解壓路徑.FullName;
                    SO = new SignOnline.SignOnline(xdoc, 簽核人員, 異動類別);
                    if (次位簽核人員 != null)
                        SO.Receiver = 次位簽核人員;
                    SO.SignerComment = 簽核人員意見;
                }
            }
            
        }


        /// <summary>
        /// 產生來文文件夾
        /// </summary>
        /// <param name="_來文文號"></param>
        /// <param name="_來文類型">函,令,開會通知單....</param>
        /// <param name="_來文檔名">電子傳di檔名紙本傳掃描tif檔名</param>
        /// <param name="_來文附件檔名">紙本傳掃描tif檔名,沒有附件請傳null</param>
        /// <returns>成功回傳空字串</returns>
        public string 產生來文文件夾(string _來文文號, string _來文類型, string _來文檔名, List<string> _來文附件檔名)
        {
            try
            {
                SO.geneRcvDocFolder(_來文文號, _來文類型, _來文檔名, _來文附件檔名);
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        /// <summary>
        /// 產生來文文件夾
        /// </summary>
        /// <param name="FepHeaderXml">由DCPWebService取得的XMLNode</param>
        /// <returns>成功回傳空字串</returns>
        public string 產生來文文件夾(XmlNode FepHeaderXml)
        {
            try
            {
                string rtn = "";
                if (FepHeaderXml.HasChildNodes)
                {
                    List<string> att = new List<string>();
                    if (FepHeaderXml.FirstChild.Name == "FepHeader")
                    {
                        foreach (XmlNode chi in FepHeaderXml.FirstChild.ChildNodes)
                        {
                            if (chi.Attributes["AtaFile"]!=null)
                                att.Add(chi.Attributes["AtaFile"].Value);
                        }
                        rtn = 產生來文文件夾(FepHeaderXml.FirstChild.Attributes["OrgDocNO"].Value, FepHeaderXml.FirstChild.Attributes["DocTypNam"].Value, FepHeaderXml.FirstChild.Attributes["DiFile"].Value, att);
                    }
                    else if (FepHeaderXml.FirstChild.Name == "ScanLog")
                    {
                        rtn = 產生來文文件夾(_文號, FepHeaderXml.FirstChild.Attributes["DocTypNam"].Value, FepHeaderXml.FirstChild.InnerText, att); 
                    }
                }
                return rtn;
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        public string 產生來文參考()
        {
            try
            {
                SO.geneRcvDocFolder();
                return "";
            }
            catch (Exception err)
            {
                return "產生來文參考時失敗!" + Environment.NewLine + err.Message;
            }
        }

        /// <summary>
        /// 加入新文書
        /// </summary>
        /// <param name="_文書名稱">簡易描述</param>
        /// <param name="_文書類型">簽,函,令,開會通知單....</param>
        /// <param name="_文書檔檔名">文書di檔</param>
        /// <param name="_文稿附件檔名">沒有附件請傳null</param>
        /// <param name="__已決">該文書是否已決行</param>
        /// <returns></returns>
        public string 加入文書封裝(string _文書名稱, string _文書類型, string _文書本文檔檔名, List<string> _文稿附件檔名,bool _已決行)
        {
            try
            {
                SO.AddWordDoc(_文書名稱, _文書類型, _文書本文檔檔名, _文稿附件檔名, _已決行);
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        /// <summary>
        /// 加入新文書
        /// </summary>
        /// <param name="_文書名稱">簡易描述</param>
        /// <param name="_文書類型">簽,函,令,開會通知單....</param>
        /// <param name="_文書檔檔名">文書di檔</param>
        /// <param name="_文稿附件檔名">沒有附件請傳null</param>
        /// <returns></returns>
        public string 加入文書封裝(string _文書名稱, string _文書類型, string _文書本文檔檔名, List<string> _文稿附件檔名)
        {
            try
            {
                SO.AddWordDoc(_文書名稱, _文書類型, _文書本文檔檔名, _文稿附件檔名);
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        /// <summary>
        /// 將該流程內容加入封裝檔並加簽
        /// </summary>
        /// <param name="憑證密碼"></param>
        /// <param name="是否轉頁面">default false</param>
        /// <returns></returns>
        public string 異動封裝檔(string 憑證密碼,bool 是否轉頁面)
        {
            SO.PKCS11Driver = this._PKCS11Driver;
            SO.PinCode = 憑證密碼;
            SO.TransdocImg = 是否轉頁面;
            SO.UseTifPageImage = _UseTifAsPage;
            SO.X509str = _CertString;
            try
            {
                SO.updateDCPFile("");
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        public string 壓縮封裝檔(string 封裝檔位置, string 輸出路徑)
        {
            try
            {
                FileInfo finfo = new FileInfo(封裝檔位置);
                SO.compressDCPFiles(finfo, 輸出路徑);
                return "";
            }
            catch (Exception err)
            {
                return "壓縮檔案發生錯誤：" + err.Message;
            }
        }

        public bool 解壓封裝檔(string 路徑與檔名, string 解壓後的路徑)
        {
            try
            {
                FileInfo fi = new FileInfo(路徑與檔名);
                GZip.Decompress(fi.DirectoryName, 解壓後的路徑, fi.Name); 
                //SO.deCompressDCPFiles(路徑與檔名, 解壓後的路徑);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 合併並會資料
        /// </summary>
        /// <param name="depSi">並會單位的封裝檔路徑</param>
        /// <returns></returns>
        public string 合併並會資料(FileInfo[] depSi)
        {
            try
            {
                foreach (FileInfo fSi in depSi)
                {
                    XmlDocument depXml = new XmlDocument();
                    depXml.XmlResolver = null;
                    depXml.Load(fSi.FullName);
                    depXml.PreserveWhitespace = true;
                    for (int i = 0; i < depXml.DocumentElement.SelectNodes("//線上簽核流程/簽核流程").Count; i++)
                    {
                        XmlNode xnd = depXml.DocumentElement.SelectNodes("//線上簽核流程/簽核流程")[i];
                        string _id = xnd.Attributes["Id"].Value;
                        XmlNode chind = null;
                        if (SO.線上簽核.DocumentElement.SelectSingleNode("//線上簽核流程/簽核流程[@Id='" + _id + "']") == null)
                        {
                            SO.線上簽核.DocumentElement.GetElementsByTagName("線上簽核流程")[0].AppendChild(SO.線上簽核.ImportNode(xnd,true));
                            chind = depXml.DocumentElement.SelectSingleNode("//線上簽核資訊/簽核點定義[@URI='#" + _id + "']");
                            SO.線上簽核.DocumentElement.GetElementsByTagName("線上簽核資訊")[0].AppendChild(SO.線上簽核.ImportNode(chind,true));
                        }
                        if (i == depXml.DocumentElement.SelectNodes("//線上簽核流程/簽核流程").Count - 1 && chind != null)
                        {
                            if (!SO.SignRefId.Contains("#" + chind.Attributes["Id"].Value))
                                SO.SignRefId.Add("#" + chind.Attributes["Id"].Value);
                        }
                    }
                }

                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        /// <summary>
        /// 補簽
        /// </summary>
        /// <param name="SignId">預補簽的簽核點定義節點Id</param>
        /// <param name="pwd">憑證密碼</param>
        /// <param name="reason">補簽原因</param>
        /// <returns></returns>
        public string 補簽(string SignId, string pwd, string reason)
        {
            string resigntime = geneTime.getTimeNow();
            SO.PKCS11Driver = this._PKCS11Driver;
            return SO.ReSignXml(SignId, pwd, reason, resigntime);
        }

        public void 設定彙併辦(string 母文文號, string[] 子文文號)
        {
            SO.SetMergeDoc(母文文號, 子文文號);
        }

    }
}
