using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.ReceivedDoc;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef;
using DigitalSealed.AttachmentsDesc;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature;
using System.IO;
using System.Collections;
using System.Drawing.Imaging;

namespace DigitalSealed.SignOnline
{
    /// <summary>
    /// 線上簽核
    /// </summary>
    internal class SignOnline 
    {
        
        private XmlDocument _SiXml = null;
        private FlowInfo _簽核流程 = null;
        private ReceivedDoc _來文文件夾 = null;
        private SignDocFolder _簽核文件夾 = null;
        private SignDocList _簽核文稿清單 = null;
        private FilesList _檔案清單 = null;
        private SignInfo _簽核資訊 = null;
        private ModifyInfo _異動資訊 = null;
        private ObjectTag _物件標籤 = null;
        private SignatureTag _加簽標籤 = null;
        private SignPointDef _簽核點定義 = null;
        private int _原始檔序號 = 0;
        private string _簽章資訊檔 = "DcpSignInfo.htm";
        private string _signTagName = "簽核文件夾";
        private string _母文文號 = "";
        private List<string> _子文文號 = null;
        private bool IsResign = false;

        #region 屬性
        public XmlDocument 線上簽核
        {
            get { return _SiXml; }
            set { _SiXml = value; }
        }

        private string _DocNO = "";
        /// <summary>
        /// 公文文號
        /// </summary>
        public string DocNO
        {
            get { return _DocNO; }
            set { _DocNO = value; }
        }

        private ModifyInfo.signer _receiver = null;
        /// <summary>
        /// 次位簽核人員
        /// </summary>
        public ModifyInfo.signer Receiver
        {
            get { return _receiver; }
            set { _receiver = value; }
        }

        private ModifyInfo.signer _sender = null;
        /// <summary>
        /// 簽核人員
        /// </summary>
        public ModifyInfo.signer Sender
        {
            get { return _sender; }
            set { _sender = value; }
        }

        private string _pinCode = "";
        /// <summary>
        /// 憑證密碼
        /// </summary>
        public string PinCode
        {
            get { return _pinCode; }
            set { _pinCode = value; }
        }

        private string _signerComment = "";
        /// <summary>
        /// 簽核人員意見
        /// </summary>
        public string SignerComment
        {
            get { return _signerComment; }
            set { _signerComment = value; }
        }

        private bool _transdocImg = false;
        /// <summary>
        /// 是否轉頁面影像檔
        /// </summary>
        public bool TransdocImg
        {
            get { return _transdocImg; }
            set { _transdocImg = value; }
        }

        /// <summary>
        /// 簽核的referenceID
        /// </summary>
        private List<string> _SignRefId = new List<string>();

        public List<string> SignRefId
        {
            get { return _SignRefId; }
            set { _SignRefId = value; }
        }

        /// <summary>
        /// 傳回目前流程點ID
        /// </summary>
        public string SignID
        {
            get
            {
                if (_簽核流程 != null)
                    return _簽核流程.Id;
                else
                    return "";
            }
        }

        private string _PKCS11Driver = "";

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
            set { _HashAlgorithm = value; }
        }

        private string _X509str = "";
        /// <summary>
        /// 憑證字串
        /// </summary>
        public string X509str
        {
            get { return _X509str; }
            set { _X509str = value; }
        }

        private bool _useTifPageImage = true;
        /// <summary>
        /// 使用單一tif描述頁面資訊，false的話使用jpg。
        /// </summary>
        public bool UseTifPageImage
        {
            get { return _useTifPageImage; }
            set { _useTifPageImage = value; }
        }

        #endregion

        #region 建構子
        public SignOnline()
        {
        }
        /// <summary>
        /// 線上簽核建構子
        /// </summary>
        /// <param name="worddocno">文書號</param>
        public SignOnline(string docno, ModifyInfo.signer sender, FlowInfo.FlowType actType)
        {
            _DocNO = docno;
            createSiXml();
            _簽核流程 = new FlowInfo(sender, actType);
            _sender = sender;
        }

        /// <summary>
        /// 線上簽核建構子
        /// </summary>
        /// <param name="siXml">帶入已經產生的線上簽核檔xml</param>
        public SignOnline(XmlDocument siXml, ModifyInfo.signer sender, FlowInfo.FlowType actType)
        {
            _SiXml = siXml;
            _DocNO = _SiXml.DocumentElement.Attributes["Id"].Value.Split(new char[] { '_' })[1];
            string 簽核流程Id = "Flow_" + sender.單位代碼 + "_" + sender.姓名 + "_" + actType.ToString();

            //先判斷如果最後一個簽核點是自己就先移除再新增避免每次存檔都產生簽核點定義，並會須避開此判斷
            if (actType !=  FlowInfo.FlowType.並會整併)
                chkIsRecentPoint(簽核流程Id);

            _簽核流程 = new FlowInfo(get簽核流程Id(簽核流程Id, actType.ToString()), actType);
            int osel = count原始檔序號();
            _原始檔序號 = osel == 0 ? 0 : osel + 1;
            _sender = sender;
            //if (actType != FlowInfo.FlowType.補簽)
            //    geneRcvDocFolder();//建立來文參考文件夾
        }

        #endregion

        private string get簽核流程Id(string 簽核流程Id,string actType)
        {
            //string 簽核流程Id = "Flow_" + sender.單位代碼 + "_" + sender.姓名 + "_" + modename.ToString();
            XmlNodeList xnls = _SiXml.DocumentElement.SelectNodes("//線上簽核流程/簽核流程[contains(@Id ,'" + 簽核流程Id + "')]");
            
            if (xnls.Count > 0)
            {
                int big = 0;
                foreach (XmlNode x in xnls)
                {
                    if (actType == FlowInfo.FlowType.並會.ToString() && x.Attributes["Id"].Value.IndexOf(FlowInfo.FlowType.並會整併.ToString()) > -1) continue;
                    big = (big > int.Parse(x.Attributes["Id"].Value.Replace(簽核流程Id, "0"))) ? big : int.Parse(x.Attributes["Id"].Value.Replace(簽核流程Id, "0"));  
                }
                
                簽核流程Id += (big + 1).ToString(); 
            }
            return 簽核流程Id;
        }

        /// <summary>
        /// 先判斷如果最後一個簽核點是自己就先移除再新增避免每次存檔都產生簽核點定義
        /// </summary>
        /// <param name="flowid">簽核流程Id</param>
        private void chkIsRecentPoint(string flowid)
        {
            XmlNodeList lastflow = _SiXml.DocumentElement.GetElementsByTagName("簽核流程");
            if (lastflow.Count > 0)
            {
                XmlNode lastflownd = lastflow[lastflow.Count-1];
                if (lastflownd.Attributes["Id"].Value.StartsWith(flowid)) _SiXml.DocumentElement.GetElementsByTagName("線上簽核流程")[0].RemoveChild(lastflownd);
            }

            XmlNodeList lastsign = _SiXml.DocumentElement.GetElementsByTagName("簽核點定義");
            if (lastsign.Count > 0)
            {
                XmlNode lastsignnd = lastsign[lastsign.Count - 1];
                string signid = "sign_" + flowid;
                if (lastsignnd.Attributes["Id"].Value.StartsWith(signid)) _SiXml.DocumentElement.GetElementsByTagName("線上簽核資訊")[0].RemoveChild(lastsignnd);
            }
        }

        /// <summary>
        /// 建立空的線上簽核XML
        /// </summary>
        private void createSiXml()
        {
            XmlDocument siXml = new XmlDocument();
            siXml.XmlResolver = null;
            siXml.AppendChild(siXml.CreateXmlDeclaration("1.0", "utf-8", null));
            siXml.AppendChild(siXml.CreateDocumentType("線上簽核", null, "99_sign_utf8.dtd", null));
             
            XmlNode root = xmlTool.MakeNode("線上簽核", "Id", "sign_" + _DocNO);
            root.AppendChild(xmlTool.MakeNode("線上簽核流程","Id","FlowInfo"));
            root.AppendChild(xmlTool.MakeNode("線上簽核資訊", "Id", "SignInfo"));
            siXml.AppendChild(siXml.ImportNode(root, true));
            siXml.PreserveWhitespace = true;
            _SiXml = siXml;
        }

        /// <summary>
        /// 取得目前流程中最大的原始檔案序號
        /// </summary>
        /// <returns></returns>
        private int count原始檔序號()
        {
            if (_SiXml != null)
            {
                XmlNodeList tmpNode = _SiXml.DocumentElement.SelectNodes("//檔案清單/電子檔案資訊[not(preceding-sibling::電子檔案資訊/@原始檔序號 > @原始檔序號 or following-sibling::電子檔案資訊/@原始檔序號 > @原始檔序號)][last()]");
                if (tmpNode.Count == 0)
                    return 0;
                else
                    return Convert.ToInt16(tmpNode[tmpNode.Count -1].Attributes["原始檔序號"].Value);
            }
            return -1;
        }

        private string get原始檔序號()
        {
            string rtn = _原始檔序號.ToString();  
            _原始檔序號++;
            return rtn;
        }

        private string get簽核文件夾Id()
        {
            string signid = "";
            if (_SiXml != null)
            {
                int seed = 0;
                
                while (true)
                {
                    signid = "DOC_" + _簽核流程.Id + (seed == 0 ? "" : seed.ToString());
                    XmlNodeList xnls = _SiXml.DocumentElement.SelectNodes("//簽核文件夾[@Id='" + signid + "']");
                    if (xnls.Count == 0) break;
                    seed++;
                }
                
            }
            return signid;
        }

        private string getObjectId()
        {
            string signid = "";
            if (_SiXml != null)
            {
                
                int seed = 0;
                while (true)
                {
                    signid = "Sign_info_" + _簽核流程.Id + (seed == 0 ? "" : seed.ToString());
                    XmlNodeList xnls = _SiXml.DocumentElement.SelectNodes("//Object[@Id='" + signid + "']");
                    if (xnls.Count == 0) break;
                    seed++;
                }
                
            }
            return signid;
        }

        private string getInDocId()
        {
            string signid = "";
            if (_SiXml != null)
            {
                int seed = 1;
                while (true)
                {
                    signid = "DOC_" + seed.ToString("0000");
                    XmlNodeList xnls = _SiXml.DocumentElement.SelectNodes("//來文文件夾[@Id='" + signid + "']");
                    if (xnls.Count == 0) break;
                    seed++;
                }

            }
            return signid;
        }

        private string getSignatureId()
        {
            string signid = "";
            if (_SiXml != null)
            {
                
                int seed = 0;
                while (true)
                {
                    signid = "ds_" + _簽核流程.Id + (seed == 0 ? "" : seed.ToString());
                    XmlNodeList xnls = _SiXml.DocumentElement.SelectNodes("//Signature[@Id='" + signid + "']");
                    if (xnls.Count == 0) break;
                    seed++;
                }
                
            }
            return signid;
        }

        /// <summary>
        /// 取出目前最後的簽核點定義Id值
        /// </summary>
        /// <returns></returns>
        private string getLastSignPointId()
        {
            XmlNodeList xnl = _SiXml.DocumentElement.GetElementsByTagName("簽核點定義");
            if (xnl.Count > 0)
                return "#" + xnl[xnl.Count - 1].Attributes["Id"].Value;
            else
                return "";
        }

        public void updateDCPFile(string ObjectId)
        {
            //geneTime.getCounter("",_DocNO); 
            if (_簽核流程.ModeName == FlowInfo.FlowType.補簽)
            {
                throw new Exception("補簽機制無法異動封裝檔!!");
            }
            //更新線上簽核流程
            updateOnlineFlowInfo();
            //geneTime.getCounter("產生線上簽核流程"); 
            //更新線上簽核資訊
            updateOnlineSignInfo();
            //geneTime.getCounter("產生線上簽核資訊"); 
            //執行簽核點加簽
            _SiXml = SignXmlSignature(_SiXml, _加簽標籤.Id, _pinCode);
            //geneTime.getCounter("End"); 
        }

        #region XML加簽
        private XmlDocument SignXmlSignature(XmlDocument xmldata,string SigId,string pwd)
        {
            Signature sig = new Signature(xmldata);
            sig.XmlFileName = _DocNO + ".si";
            sig.CertX509Str = _X509str;
            if (_PKCS11Driver != "") sig.PKCS11Driver = _PKCS11Driver;
            sig.FilePath = ConstVar.SourceFilesPath;
            if (SigId == "")
                SigId = _加簽標籤.Id;
            //geneTime.getCounter("加簽類別建構子");
            if (IsResign)
                return sig.getReSignXmlData(xmldata.InnerXml, SigId, pwd);
            else
                return sig.getSignXmlData(SigId, pwd);
        }

        /// <summary>
        /// 補簽
        /// </summary>
        /// <param name="SignId">補簽的簽核點定義Id</param>
        /// <param name="pwd">憑證密碼</param>
        /// <param name="reason">補簽原因</param>
        /// <param name="resigntime">補簽時間</param>
        internal string ReSignXml(string SignId, string pwd,string reason,string resigntime)
        {
            try
            {
                
                string SignatureId = SignId.Replace("sign_", "ds_");
                if (SignatureId.StartsWith("Flow_")) SignatureId = "ds_" + SignatureId;
                //重簽
                IsResign = true;
                XmlDocument redoc = SignXmlSignature(_SiXml, SignatureId, pwd);
                IsResign = false;
                XmlNode ReSignatureNd = redoc.DocumentElement.SelectSingleNode("//Signature[@Id='" + SignatureId + "']");
                ReSignatureNd.Attributes["Id"].Value = SignatureId + "_補簽";
                //判斷是否已有補簽追認的節點
                XmlNodeList resignNode = _SiXml.DocumentElement.GetElementsByTagName("補簽追認");
                if (resignNode.Count == 0)
                {
                    XmlNode _補簽追認 = xmlTool.MakeNode("補簽追認", "補簽人員數", "1");
                    Dictionary<string, string> attb = new Dictionary<string, string>();
                    attb.Add("URI", "#" + SignId);
                    attb.Add("補簽註記", reason);
                    attb.Add("產生時間", resigntime);
                    XmlNode _補簽人員 = xmlTool.MakeNode("補簽人員", attb);
                    XmlNode _簽章資料 = xmlTool.MakeNode("簽章資料", "");
                    _簽章資料.AppendChild(_簽章資料.OwnerDocument.ImportNode(_sender.getSignerNodeXML(), true));
                    _簽章資料.AppendChild(_簽章資料.OwnerDocument.ImportNode(ReSignatureNd, true));
                    _補簽人員.AppendChild(_補簽人員.OwnerDocument.ImportNode(_簽章資料, true));
                    _補簽追認.AppendChild(_補簽追認.OwnerDocument.ImportNode(_補簽人員, true));
                    _SiXml.DocumentElement.GetElementsByTagName("線上簽核資訊")[0].AppendChild(_SiXml.ImportNode( _補簽追認,true));
                }
                else
                {
                    XmlNode _補簽追認 = resignNode[0];
                    _補簽追認.Attributes["補簽人員數"].Value = Convert.ToString(int.Parse(_補簽追認.Attributes["補簽人員數"].Value) + 1);
                    Dictionary<string, string> attb = new Dictionary<string, string>();
                    attb.Add("URI", "#" + SignId);
                    attb.Add("補簽註記", reason);
                    attb.Add("產生時間", resigntime);
                    XmlNode _補簽人員 = xmlTool.MakeNode("補簽人員", attb);
                    XmlNode _簽章資料 = xmlTool.MakeNode("簽章資料", "");
                    _簽章資料.AppendChild(_簽章資料.OwnerDocument.ImportNode(_sender.getSignerNodeXML(), true));
                    _簽章資料.AppendChild(_簽章資料.OwnerDocument.ImportNode(ReSignatureNd, true));
                    _補簽人員.AppendChild(_補簽人員.OwnerDocument.ImportNode(_簽章資料, true));
                    _補簽追認.AppendChild(_補簽追認.OwnerDocument.ImportNode(_補簽人員, true));
                }
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
            
        }
        #endregion

        #region 線上簽核流程

        private void updateOnlineFlowInfo()
        {
            XmlNodeList xnls = _SiXml.DocumentElement.GetElementsByTagName("線上簽核流程");
            if (xnls.Count == 1)
            {
                xnls[0].AppendChild(_SiXml.ImportNode( _簽核流程.getFlowInfoNode(),true));
            }
        }

        #endregion

        #region 線上簽核資訊

        private void updateOnlineSignInfo()
        {
            geneSignPointDef();
            XmlNodeList xnls = _SiXml.DocumentElement.GetElementsByTagName("線上簽核資訊");
            if (xnls.Count == 1)
            {
                xnls[0].AppendChild(_SiXml.ImportNode(_簽核點定義.getSignPointNode(),true));    
            }
        }

        #region 簽核點定義

        private void geneSignPointDef()
        {
            geneObjectTag();
            geneSignatureTag();
            string preid = getLastSignPointId();
            List<string> preSignPointId = new List<string>();
            if (preid != "")
            {
                preSignPointId.Add(getLastSignPointId());
            }
            _簽核點定義 = new SignPointDef(_簽核流程.Id, preSignPointId, _加簽標籤, _物件標籤);
            if (_SignRefId.Count > 0)
            {
                foreach (string _id in _SignRefId)
                {
                    if (!_簽核點定義.RefSignDefId.Contains(_id) )
                        _簽核點定義.RefSignDefId.Add(_id);   
                }
                
            }
        }

        #region SignatureTag

        private void geneSignatureTag()
        {
            _加簽標籤 = new SignatureTag(getSignatureId(), _物件標籤);
            _加簽標籤.HashAlgorithm = _HashAlgorithm;
        }

        #endregion

        #region ObjectTag

        private void geneObjectTag()
        {
            geneModifyInfo();
            geneSignInfo();
            _物件標籤 = new ObjectTag(getObjectId(), _異動資訊, _簽核資訊);   
        }

        #region 異動資訊
        private void geneModifyInfo()
        {
            if (_receiver != null)
                _異動資訊 = new ModifyInfo(_sender, _receiver, _簽核流程.ModeName, _signerComment, geneTime.getTimeNow());
            else
                _異動資訊 = new ModifyInfo(_sender, _簽核流程.ModeName, _signerComment, geneTime.getTimeNow());
        }
        #endregion

        #region 簽核資訊

        private void geneSignInfo()
        {
            geneSignDocFolder();
            if (_來文文件夾 != null)
                _簽核資訊 = new SignInfo(_簽核文件夾, _來文文件夾);
            else
                _簽核資訊 = new SignInfo(_簽核文件夾);
        }

        #region 來文文件夾
        /// <summary>
        /// 產生來文文件夾
        /// </summary>
        internal void geneRcvDocFolder(string _來文文號, string _來文類型, string _來文檔名, List<string> _來文附件檔名)
        {
            string _產生時間 = geneTime.getTimeNow();
            _檔案清單 = new FilesList();
            if (_來文檔名.ToLower().EndsWith(".di"))
            {
                _來文文件夾 = new ReceivedDoc(getInDocId(), _來文文號, _來文類型, "電子");
                //來文di檔不能異動只能維持第一個簽核點的檔案資訊
                XmlNode xnd = _SiXml.DocumentElement.SelectSingleNode("//電子檔案資訊[@原始檔序號='0']");
                if (xnd!=null)
                    _檔案清單.Add(new ElecFileInfo(_來文檔名, "0", xnd.Attributes["產生時間"].Value));
                else
                    _檔案清單.Add(new ElecFileInfo(_來文檔名, get原始檔序號() , _產生時間));
            }
            else //只處理紙本掃描tif檔
            {
                if (_來文檔名.ToLower().EndsWith(".tif") || _來文檔名.ToLower().EndsWith(".tiff"))
                {
                    _來文文件夾 = new ReceivedDoc(getInDocId(), _來文文號, _來文類型, "紙本");
                    FileInfo fi = new FileInfo(Path.Combine(ConstVar.SourceFilesPath, _來文檔名));
                    string orgno = get原始檔序號();
                    _檔案清單.Add(new ElecFileInfo(fi.Name, orgno, _產生時間));
                    _來文文件夾.indoc.AddPage(new DocPageList.Page(orgno));
                    //foreach (string ono in SplitRcvDocImg(_來文檔名))      //來文影像掃描不拆一頁一檔頁面
                    //    _來文文件夾.indoc.AddPage(new DocPageList.Page(ono));
                }
                else
                    throw new Exception("來文掃描非Tif檔無法處理!");
            }
            _來文文件夾.GeneTime = _產生時間;
            //處理來文附件部分
            foreach (string fname in _來文附件檔名)
            {
                if (_來文文件夾.DocFormat == "電子")
                {
                    string orgno = get原始檔序號();
                    IAttachment iatt = new ElecAtt(fname, _產生時間, AttachStyle.電子檔格式附件, orgno);
                    _來文文件夾.indoc.AddAtt(iatt);
                    //_檔案清單.Add(new ElecFileInfo(fname, orgno, _產生時間));
                    _檔案清單.Add((ElecFileInfo)(iatt as ElecAtt));
                }
                else
                {
                    if (fname.ToLower().EndsWith(".tif") || fname.ToLower().EndsWith(".tiff"))
                    {
                        foreach (KeyValuePair<string, string> ono in SplitRcvDocAttImg(fname))
                        {
                            IAttachment iatt = new ElecAtt(ono.Value, _產生時間, AttachStyle.一般頁面檔格式附件, ono.Key);
                            _來文文件夾.indoc.AddAtt(iatt);
                        }
                    }
                    else
                        throw new Exception("紙本來文附件非Tif檔無法處理!");
                }
            }
        }


        /// <summary>
        /// 加入來文文件夾參考
        /// </summary>
        internal void geneRcvDocFolder()
        {
            XmlNodeList xnls = _SiXml.DocumentElement.GetElementsByTagName("來文文件夾");
            if (xnls.Count > 0)
            {
                _來文文件夾 = new ReceivedDoc(xnls[xnls.Count-1].Attributes["Id"].Value);
                if (_檔案清單 == null) _檔案清單 = new FilesList();
                XmlNode preFileList = xnls[xnls.Count - 1].NextSibling.SelectSingleNode("檔案清單");
                if (xnls[xnls.Count - 1].FirstChild.FirstChild.Attributes["格式"].Value == "電子")
                {
                    XmlNode preInDOC = preFileList.SelectSingleNode("//電子檔案資訊[@原始檔序號='0']");
                    if (preInDOC == null) {throw new Exception("si檔缺少來文di檔的描述");}
                    _檔案清單.Add(new ElecFileInfo(preInDOC.FirstChild.InnerText, "0", preInDOC.Attributes["產生時間"].Value));
                }
                else
                {
                    XmlNodeList 來文頁面 = xnls[xnls.Count - 1].SelectNodes("//來文清單/來文/紙本來文/頁面[@原始檔序號]");
                    for (int j = 0; j < 來文頁面.Count; j++)
                    {
                        string cnt = 來文頁面[j].Attributes["原始檔序號"].Value;
                        XmlNode preInAtt = preFileList.SelectSingleNode("//電子檔案資訊[@原始檔序號='" + cnt + "']");
                        if (preInAtt == null) {throw new Exception("si檔缺少紙本來文檔的描述");}
                        _檔案清單.Add(new ElecFileInfo(preInAtt.FirstChild.InnerText, cnt, preInAtt.Attributes["產生時間"].Value));        
                    }
                }

                
                //處理來文附件
                XmlNodeList orgno = xnls[xnls.Count - 1].SelectNodes("//來文清單/來文/*/*/*/*[@原始檔序號]");
                for (int j = 0; j < orgno.Count; j++)
                {
                    string cnt = orgno[j].Attributes["原始檔序號"].Value;
                    XmlNode preInAtt = preFileList.SelectSingleNode("//電子檔案資訊[@原始檔序號='" + cnt + "']");
                    if (preInAtt == null) {throw new Exception("來文附件描述有誤!找不到原始檔序號=" + cnt.ToString());}
                    _檔案清單.Add(new ElecFileInfo(preInAtt.FirstChild.InnerText, cnt, preInAtt.Attributes["產生時間"].Value));        
                }
            }
        }

        #endregion

        #region 簽核文件夾

        private void geneSignDocFolder()
        {
            if (_檔案清單 == null) _檔案清單 = new FilesList();  

            //將文稿資料轉入檔案清單中
            if (_簽核文稿清單 != null)
            {
                foreach (Document d in _簽核文稿清單)
                {
                    if (_transdocImg)
                        transDocToImg(d, _useTifPageImage);
                    _檔案清單.Add(new ElecFileInfo(d.DFilename, d.Ofileserial, d.GeneTime));
                    if (d.附件清單 != null)
                    {
                        foreach (ElecAtt ea in d.附件清單)
                        {
                            _檔案清單.Add((ElecFileInfo)ea);
                        }
                    }
                }
            }
            _簽核文件夾 = new SignDocFolder(get簽核文件夾Id(), geneTime.getTimeNow(), _簽核文稿清單, _檔案清單);
            if (_簽核流程.ModeName == FlowInfo.FlowType.解併)
                _signTagName = "子文簽核文面夾";
            _簽核文件夾.TagName = _signTagName;
            _簽核文件夾.併文清單 = new MergeFileList(_母文文號, _子文文號);
        }

        /// <summary>
        /// 依據文書決行狀態轉頁面
        /// </summary>
        /// <param name="byOEJudge"></param>
        private void geneSignDocFolder(bool byOEJudge)
        {
            if (_檔案清單 == null) _檔案清單 = new FilesList();

            //將文稿資料轉入檔案清單中
            if (_簽核文稿清單 != null)
            {
                foreach (Document d in _簽核文稿清單)
                {
                    if (d.Judged)
                        transDocToImg(d, true);
                    _檔案清單.Add(new ElecFileInfo(d.DFilename, d.Ofileserial, d.GeneTime));
                    if (d.附件清單 != null)
                    {
                        foreach (ElecAtt ea in d.附件清單)
                        {
                            _檔案清單.Add((ElecFileInfo)ea);
                        }
                    }
                }
            }
            _簽核文件夾 = new SignDocFolder(get簽核文件夾Id(), geneTime.getTimeNow(), _簽核文稿清單, _檔案清單);
            _簽核文件夾.TagName = _signTagName;
            _簽核文件夾.併文清單 = new MergeFileList(_母文文號, _子文文號);
        }


        internal void SetMergeDoc(string 母文文號, string[] 子文文號)
        {
            _子文文號 = new List<string>(子文文號);
            _母文文號 = 母文文號;
            if (_子文文號.Contains(_DocNO)) _signTagName = "子文簽核文件夾";
            
        }

        #region 簽核文稿清單

        internal void AddWordDoc(string _文書名稱, string _文書類型, string _文書檔檔名, List<string> _文稿附件檔名,bool _已決行)
        {
            if (_簽核文稿清單 == null)
                _簽核文稿清單 = new SignDocList();
            string _產生時間 = geneTime.getTimeNow();
            Document _文稿 = new Document(_文書檔檔名, _文書名稱, get原始檔序號(), _產生時間, _文書類型, _已決行);
            if (_文稿附件檔名 != null)
            {
                if (_文稿附件檔名.Count > 0) _文稿.附件清單 = new AttsList();
                foreach (string s in _文稿附件檔名)
                    _文稿.附件清單.Add(new ElecAtt(s, _產生時間, AttachStyle.電子檔格式附件, get原始檔序號()));
            }
            _簽核文稿清單.Add(_文稿);
        }

        internal void AddWordDoc(string _文書名稱, string _文書類型, string _文書檔檔名, List<string> _文稿附件檔名)
        {
            if (_簽核文稿清單 == null)
                _簽核文稿清單 = new SignDocList();
            string _產生時間 = geneTime.getTimeNow();
            Document _文稿 = new Document(_文書檔檔名, _文書名稱, get原始檔序號(), _產生時間, _文書類型);
            if (_文稿附件檔名 != null)
            {
                if (_文稿附件檔名.Count > 0) _文稿.附件清單 = new AttsList();
                foreach (string s in _文稿附件檔名)
                    _文稿.附件清單.Add(new ElecAtt(s, _產生時間, AttachStyle.電子檔格式附件, get原始檔序號()));
            }
            _簽核文稿清單.Add(_文稿);
        }

        #endregion
        #endregion
        #endregion
        #endregion
        #endregion
        #endregion

        #region 封裝檔案壓縮解壓縮

        public void compressDCPFiles(string siFilePath, string outputpath)
        {
            List<FileInfo> fifo = new List<FileInfo>();
            foreach (ElecFileInfo doc in _檔案清單)
            {
                string fullfilename = Path.Combine(ConstVar.SourceFilesPath, doc.Filename);
                if (File.Exists(fullfilename))
                {
                    FileInfo fi = new FileInfo(fullfilename);
                    fifo.Add(fi); 
                }
            }
            if (File.Exists(siFilePath))
            {
                FileInfo fi = new FileInfo(siFilePath);
                fifo.Add(fi);
            }
            if (outputpath == "") outputpath = ConstVar.SourceFilesPath;
            GZip.Compress(fifo, fifo[0].DirectoryName, outputpath, _DocNO + ".dsicdp");      
        }

        public void compressDCPFiles(FileInfo siFilePath, string outputpath)
        {
            if (outputpath == "") outputpath = ConstVar.SourceFilesPath;
            GZip.Compress(siFilePath.Directory.GetFiles(), siFilePath.DirectoryName, outputpath, _DocNO + ".dsicdp");
        }

        private void deCompressDCPFiles(string FullGZFilePath,string OutputFilesPath)
        {
            FileInfo fi = new FileInfo(FullGZFilePath);
            GZip.Decompress(fi.DirectoryName, OutputFilesPath, fi.Name);     
        }

        #endregion

        #region  文稿轉頁面影像檔
        private void transDocToImg(Document aDoc)
        {
            FileInfo fi = new FileInfo(Path.Combine(ConstVar.SourceFilesPath,aDoc.DFilename));
            if (fi.Exists)
            {
                string pagefile = fi.FullName.Replace(fi.Extension, ".tif");  //fi.FullName.Substring(0, fi.FullName.LastIndexOf("V")) + ".tif";
                string VerStr = fi.FullName.Substring(fi.FullName.LastIndexOf("V"), fi.FullName.LastIndexOf(".") - fi.FullName.LastIndexOf("V"));
                pagefile = pagefile.Replace(VerStr + ".tif", ".tif"); 
                if (File.Exists(pagefile))
                {
                    DirectoryInfo dir = Directory.CreateDirectory(Path.Combine(ConstVar.FileWorkingPath, _DocNO));
                    foreach (FileInfo f in dir.GetFiles())
                    {
                        f.Delete();
                    }

                    if (pagefile.EndsWith(".tif"))
                    {
                        try
                        {
                            ConvertTifToJpg(pagefile, dir, VerStr);
                        }
                        catch (Exception err)
                        {
                            throw new Exception("轉頁面檔失敗!" + err.Message);
                        }
                    }
                    else
                    {
                        //轉個文書頁面 來源:"文書號.htm" 目的:"簽核文件夾.Id+轉出後的預設檔名"
                        //ImageConverter ic = new ImageConverter();

                        //bool rtn = ic.do_veryPDF_docPrintCom(pagefile, Path.Combine(dir.FullName, fi.Name.Replace(fi.Extension.ToString(), ".jpg")), "");
                        //if (!rtn) throw new Exception("文書轉頁面檔失敗! " + ic.strLastErrorMessage);

                        //轉簽章資訊頁面 "文書號DcpSignInfo.htm"
                        //string signfilename = fi.FullName.Substring(0, fi.FullName.LastIndexOf("V")) + _簽章資訊檔;
                        //if (File.Exists(signfilename))
                        //{
                        //    rtn = ic.do_veryPDF_docPrintCom(Path.Combine(ConstVar.SourceFilesPath, _簽章資訊檔), Path.Combine(dir.FullName, fi.Name.Replace(fi.Extension.ToString(), "SignInfo.jpg")), "");
                        //    if (!rtn) throw new Exception("轉簽核意見頁面失敗! " + ic.strLastErrorMessage);
                        //}
                        //else
                        //    throw new Exception("找不到簽核意見的htm檔!");
                    }
                    foreach (FileInfo aFi in dir.GetFiles())
                    {
                        string targetfile = Path.Combine(ConstVar.SourceFilesPath, aFi.Name);
                        XmlNodeList nsts = this._SiXml.DocumentElement.SelectNodes("//電子檔案資訊/檔案名稱[.='" + aFi.Name + "']");
                        if (nsts.Count > 0)
                        {
                            if (File.Exists(targetfile))
                            {
                                string serialno = nsts[nsts.Count - 1].ParentNode.Attributes["原始檔序號"].Value;
                                string modetime = nsts[nsts.Count - 1].ParentNode.Attributes["產生時間"].Value;
                                aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                                _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, modetime));
                            }
                            else
                            {
                                aFi.CopyTo(targetfile, false);
                                string serialno = get原始檔序號();
                                aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                                _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, geneTime.getTimeNow()));
                            }
                        }
                        else
                        {
                            aFi.CopyTo(targetfile, true);
                            string serialno = get原始檔序號();
                            aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                            _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, geneTime.getTimeNow()));
                        }

                        //if (!File.Exists(targetfile))
                        //{
                        //    aFi.CopyTo(targetfile, false);
                        //    string serialno = get原始檔序號();
                        //    aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                        //    _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, geneTime.getTimeNow()));
                        //}
                        //else
                        //{
                        //    XmlNodeList nsts = this._SiXml.DocumentElement.SelectNodes("//電子檔案資訊/檔案名稱[.='" + aFi.Name + "']");
                        //    if (nsts.Count > 0)
                        //    {
                        //        string serialno = nsts[nsts.Count - 1].ParentNode.Attributes["原始檔序號"].Value;
                        //        string modetime = nsts[nsts.Count - 1].ParentNode.Attributes["產生時間"].Value;
                        //        aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                        //        _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, modetime));
                        //    }
                        //}
                    }
                }
                else
                {
                    //找不到要轉換的文稿htm檔
                    throw new Exception("找不到轉頁面的原始檔!");
                }

            }
        }

        /// <summary>
        /// 採一檔多頁的方式描述封裝
        /// </summary>
        /// <param name="aDoc"></param>
        /// <param name="OneTif">是否使用一檔多頁的tif</param>
        private void transDocToImg(Document aDoc,bool OneTif)
        {
            if (OneTif)
            {
                FileInfo fi = new FileInfo(Path.Combine(ConstVar.SourceFilesPath, aDoc.DFilename));
                if (fi.Exists)
                {
                    string pagefile = fi.FullName.Replace(fi.Extension, ".tif");  //fi.FullName.Substring(0, fi.FullName.LastIndexOf("V")) + ".tif";
                    string VerStr = fi.FullName.Substring(fi.FullName.LastIndexOf("V"), fi.FullName.LastIndexOf(".") - fi.FullName.LastIndexOf("V"));
                    if (File.Exists(pagefile.Replace(VerStr + ".tif", ".tif")))
                    {
                        if (File.Exists(pagefile)) File.Delete(pagefile);
                        File.Move(pagefile.Replace(VerStr + ".tif", ".tif"), pagefile);
                    }

                    if (!File.Exists(pagefile))
                        return;

                    FileInfo aFi = new FileInfo(pagefile);

                    XmlNodeList nsts = this._SiXml.DocumentElement.SelectNodes("//電子檔案資訊/檔案名稱[.='" + aFi.Name + "']");
                    if (nsts.Count > 0)
                    {
                        string serialno = nsts[nsts.Count - 1].ParentNode.Attributes["原始檔序號"].Value;
                        string modetime = nsts[nsts.Count - 1].ParentNode.Attributes["產生時間"].Value;
                        aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                        _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, modetime));
                    }
                    else
                    {
                        string serialno = get原始檔序號();
                        aDoc.文稿頁面檔.文稿頁面清單.Add(new DocPageList.Page(serialno));
                        _檔案清單.Add(new ElecFileInfo(aFi.Name, serialno, geneTime.getTimeNow()));
                    }
                }
            }
            else
                transDocToImg(aDoc);
        }

        private bool ConvertTifToJpg(string tiffile,DirectoryInfo targetfolder,string verno)
        {
            try
            {
                FileInfo file = new FileInfo(tiffile);
                System.Drawing.Image imageFile = System.Drawing.Image.FromFile(tiffile);
                System.Drawing.Imaging.FrameDimension frameDimensions = new System.Drawing.Imaging.FrameDimension(imageFile.FrameDimensionsList[0]);
                int NumberOfFrames = imageFile.GetFrameCount(frameDimensions);
                string[] paths = new string[NumberOfFrames];
                for (int intFrame = 0; intFrame < NumberOfFrames; ++intFrame)
                {
                    //imageFile.SelectActiveFrame(frameDimensions, intFrame);
                    //System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(imageFile);
                    //paths[intFrame] = Path.Combine(targetfolder.FullName, file.Name.Replace(file.Extension, "") + verno + intFrame.ToString("00") + ".jpg");
                    //bmp.Save(paths[intFrame], System.Drawing.Imaging.ImageFormat.Jpeg);
                    //bmp.Dispose();

                    imageFile.SelectActiveFrame(frameDimensions, intFrame);
                    System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(imageFile);
                    bmp.SetResolution(200, 200);
                    paths[intFrame] = Path.Combine(targetfolder.FullName, file.Name.Replace(file.Extension, "") + verno + intFrame.ToString("00") + ".jpg");
                    ImageCodecInfo info = null;
                    foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                        if (ice.FormatID == ImageFormat.Jpeg.Guid)
                            info = ice;
                    System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.Quality;
                    System.Drawing.Imaging.Encoder colordepth = System.Drawing.Imaging.Encoder.ColorDepth;
                    EncoderParameters ep = new EncoderParameters(2);
                    ep.Param[0] = new EncoderParameter(enc, 75L);
                    ep.Param[1] = new EncoderParameter(colordepth, 24L);
                    bmp.Save(paths[intFrame], info, ep);
                    bmp.Dispose();
                }
                imageFile.Dispose();
                return true;
            }
            catch (Exception err)
            {
                throw err;
                //return false;
            }
        }
        #endregion

        #region 紙本來文影像轉一頁一檔
        /// <summary>
        /// 紙本來文轉影像
        /// </summary>
        /// <param name="DocTifFileName">tif檔一檔多頁</param>
        private List<string> SplitRcvDocImg(string TifFileName)
        {
            //建出一個準備轉檔的空間
            DirectoryInfo dir = Directory.CreateDirectory(Path.Combine(ConstVar.FileWorkingPath, _DocNO));
            foreach (FileInfo f in dir.GetFiles())
            {
                f.Delete();
            }
            TiffManager tm = new TiffManager(Path.Combine(ConstVar.SourceFilesPath, TifFileName));
            ArrayList al =  tm.SplitTiffImage(dir.FullName, System.Drawing.Imaging.EncoderValue.CompressionLZW);
            string _產生時間 = geneTime.getTimeNow();
            List<string> rtn = new List<string>(); 
            foreach (string afile in al)
            {
                FileInfo fi = new FileInfo(afile);
                File.Copy(afile, Path.Combine(ConstVar.SourceFilesPath, fi.Name),true);
                File.Delete(afile); 
                string orgno = get原始檔序號();
                _檔案清單.Add(new ElecFileInfo(fi.Name, orgno , _產生時間));
                rtn.Add(orgno);
            }
            return rtn;
        }

        private Dictionary<string,string> SplitRcvDocAttImg(string TifFileName)
        {
            //建出一個準備轉檔的空間
            DirectoryInfo dir = Directory.CreateDirectory(Path.Combine(ConstVar.FileWorkingPath, _DocNO));
            foreach (FileInfo f in dir.GetFiles())
            {
                f.Delete();
            }
            TiffManager tm = new TiffManager(Path.Combine(ConstVar.SourceFilesPath, TifFileName));
            ArrayList al = tm.SplitTiffImage(dir.FullName, System.Drawing.Imaging.EncoderValue.CompressionLZW);
            string _產生時間 = geneTime.getTimeNow();
            Dictionary<string, string> rtn = new Dictionary<string,string>();
            foreach (string afile in al)
            {
                FileInfo fi = new FileInfo(afile);
                File.Copy(afile, Path.Combine(ConstVar.SourceFilesPath, fi.Name),true);
                File.Delete(afile);
                string orgno = get原始檔序號();
                _檔案清單.Add(new ElecFileInfo(fi.Name, orgno, _產生時間));
                rtn.Add(orgno, fi.Name);
            }
            return rtn;
        }
        #endregion
    }

}
