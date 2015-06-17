using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
using System.Data.SqlTypes;
using System.IO;
using System.Configuration;
using DigitalSealed;

namespace DCPWebService
{
    /// <summary>
    /// 數位封裝專用WS
    /// </summary>
    [WebService(Namespace = "http://www.dsic.com.tw/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class DCPWSE : System.Web.Services.WebService
    {
        private string SiFilePath = "";
        private string _DepId = "";

        /// <summary>
        /// 由機關交換代碼反查機關別代碼
        /// </summary>
        /// <param name="useOrgID">機關交換代碼</param>
        /// <returns></returns>
        private short getUseOrgSeqNO(string useOrgID)
        {
            string sqlstr = "select UseOrgSeqNO from UseOrganization where UseOrgID='" + useOrgID + "'";
            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.Char,20)};
            Parms[0].Value = useOrgID;
            short rtn = 0;
            using (SqlDataReader drObj = SqlHelper.ExecuteReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                if (drObj.HasRows)
                {
                    if (drObj.Read())
                    {
                        rtn = Convert.ToInt16(drObj["UseOrgSeqNO"]);
                    }
                }
            }
            return rtn;
        }

        /// <summary>
        /// 取得公文相關文書及附件
        /// </summary>
        /// <param name="_useorgseqno">機關代碼</param>
        /// <param name="_docno">文號</param>
        /// <returns></returns>
        [WebMethod]
        public XmlNode getWordDocAttInfo(string _orgid, string _docno)
        {
            string _useorgseqno = Convert.ToString(getUseOrgSeqNO(_orgid));
            string sqlstr = "SELECT  OEHeader.OEDocNO, OEHeader.DocTypNam, OEHeader.VerNO,OEAtaDetail.AtaFile, OEAtaDetail.LastChangeVer, OEAtaDetail.AtaTyp ";
            sqlstr += "FROM OEHeader LEFT OUTER JOIN OEAtaDetail ON (OEHeader.UseOrgSeqNO=OEAtaDetail.UseOrgSeqNO and OEHeader.OEDocNO = OEAtaDetail.OEDocNO and OEAtaDetail.AtaTyp='2') WHERE (OEHeader.ParDocNO = @DocNO) AND (OEHeader.UseOrgSeqNO = @SeqNO) FOR XML AUTO";
            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt,2),
					new SqlParameter("@DocNO", SqlDbType.Char, _docno.Length)};
            Parms[0].Value = _useorgseqno;
            Parms[1].Value = _docno;
            
            using (XmlReader xdr = SqlHelper.ExecuteXmlReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                XmlNode rtnNode = new XmlDocument().CreateElement("Result") as XmlNode;
                if (xdr.Read())
                {
                    string tmp = "";
                    while (xdr.ReadState != ReadState.EndOfFile)
                    {
                        tmp += xdr.ReadOuterXml();
                    }
                    rtnNode.InnerXml = tmp;
                }
                return rtnNode;
            }
            return null;

        }

        /// <summary>
        /// 取得文書附件與決行狀態
        /// </summary>
        /// <param name="_orgid">機關代碼</param>
        /// <param name="_docno">文號</param>
        /// <returns></returns>
        [WebMethod]
        public XmlNode getWordAttJudgeInfo(string _orgid, string _docno)
        {
            string _useorgseqno = Convert.ToString(getUseOrgSeqNO(_orgid));
            string sqlstr = "SELECT OEHeader.OEDocNO, OEHeader.DocTypNam, OEHeader.VerNO, JudgeDetail.JudStampNam, OEAtaDetail.AtaFile, OEAtaDetail.LastChangeVer, OEAtaDetail.AtaTyp ";
            sqlstr += "FROM OEHeader LEFT OUTER JOIN OEAtaDetail ON (OEHeader.UseOrgSeqNO=OEAtaDetail.UseOrgSeqNO and OEHeader.OEDocNO = OEAtaDetail.OEDocNO and OEAtaDetail.AtaTyp='2') ";
            sqlstr += "LEFT OUTER JOIN JudgeDetail ON (OEHeader.UseOrgSeqNO=JudgeDetail.UseOrgSeqNO and OEHeader.OEDocNO = JudgeDetail.OEDocNO) WHERE (OEHeader.ParDocNO = @DocNO) AND (OEHeader.UseOrgSeqNO = @SeqNO) FOR XML AUTO";
            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt,2),
					new SqlParameter("@DocNO", SqlDbType.Char, _docno.Length)};
            Parms[0].Value = _useorgseqno;
            Parms[1].Value = _docno;

            using (XmlReader xdr = SqlHelper.ExecuteXmlReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                XmlNode rtnNode = new XmlDocument().CreateElement("Result") as XmlNode;
                if (xdr.Read())
                {
                    string tmp = "";
                    while (xdr.ReadState != ReadState.EndOfFile)
                    {
                        tmp += xdr.ReadOuterXml();
                    }
                    rtnNode.InnerXml = tmp;
                }
                return rtnNode;
            }
            return null;

        }

        [WebMethod]
        public XmlNode getDocRefInfo(string _orgid, string _oedocno)
        {
            string _useorgseqno = Convert.ToString(getUseOrgSeqNO(_orgid));
            string sqlstr = "SELECT  FileNam FROM DocAtaDetail WHERE (FileTyp = '2' AND UploadTime <>'' AND UseOrgSeqNO = @SeqNO AND RsvFieldA = @OEDocNO) FOR XML AUTO";
            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt,2),
					new SqlParameter("@OEDocNO", SqlDbType.Char, _oedocno.Length)};
            Parms[0].Value = _useorgseqno;
            Parms[1].Value = _oedocno;

            using (XmlReader xdr = SqlHelper.ExecuteXmlReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                XmlNode rtnNode = new XmlDocument().CreateElement("Result") as XmlNode;
                if (xdr.Read())
                {
                    string tmp = "";
                    while (xdr.ReadState != ReadState.EndOfFile)
                    {
                        tmp += xdr.ReadOuterXml();
                    }
                    rtnNode.InnerXml = tmp;
                }
                return rtnNode;
            }
            return null;

        }

        /// <summary>
        /// 取得來文相關資訊
        /// </summary>
        /// <param name="_useorgseqno">機關代碼</param>
        /// <param name="_docno">文號</param>
        /// <returns></returns>
        [WebMethod]
        public XmlNode getInDocAttInfo(string _orgid, string _docno)
        {
            string _useorgseqno = Convert.ToString(getUseOrgSeqNO(_orgid));
            string sql = "select FepRcvSeqNO from DocHeader where DocNO='" + _docno + "' and UseOrgSeqNO = " + _useorgseqno ;
            string sqlstr = "SELECT FepHeader.FepSeqNO ,FepHeader.OrgDocNO ,FepHeader.DocTypNam , FepHeader.DiFile ,FepAtaDetail.AtaFile FROM FepHeader LEFT OUTER JOIN FepAtaDetail ON ";
            sqlstr += "FepHeader.UseOrgSeqNO = FepAtaDetail.UseOrgSeqNO and FepHeader.FepSeqNO = FepAtaDetail.FepSeqNO and FepHeader.RcvOrgID = '" + _orgid + "' where FepHeader.FepSeqNO = (" + sql + ") and FepHeader.DocNO=@DocNO and FepAtaDetail.RcvOrgID='" + _orgid + "' FOR XML AUTO";
            SqlParameter[] Parms = new SqlParameter[] {
                     new SqlParameter("@SeqNO", SqlDbType.SmallInt,2),
					new SqlParameter("@DocNO", SqlDbType.Char, _docno.Length)};
            Parms[0].Value = _useorgseqno;
            Parms[1].Value = _docno;
            using (XmlReader xdr = SqlHelper.ExecuteXmlReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                XmlNode rtnNode = new XmlDocument().CreateElement("Result") as XmlNode;
                if (xdr.Read())
                {
                    string tmp = "";
                    while (xdr.ReadState != ReadState.EndOfFile)
                    {
                        tmp += xdr.ReadOuterXml();
                    }
                    rtnNode.InnerXml = tmp;
                }
                else //紙本附件
                {
                    sqlstr = "SELECT ScanLog.ImgPage, ScanLog.ImgDir, DocHeader.DocTypNam FROM  ScanLog left join DocHeader on ScanLog.UseOrgSeqNo = DocHeader.UseOrgSeqNo and  ScanLog.DocNO = DocHeader.DocNO WHERE ScanLog.UseOrgSeqNo=@SeqNO AND ScanLog.DocNO = @DocNO";
                    using (SqlDataReader sdr = SqlHelper.ExecuteReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
                    {
                        while (sdr.Read())
                        {
                            XmlNode anode = rtnNode.OwnerDocument.CreateElement("ScanLog");
                            XmlAttribute att1 = rtnNode.OwnerDocument.CreateAttribute("PageCnt");
                            att1.Value = sdr["ImgPage"].ToString();
                            XmlAttribute att2 = rtnNode.OwnerDocument.CreateAttribute("ImgDir");
                            att2.Value = sdr["ImgDir"].ToString();
                            XmlAttribute att3 = rtnNode.OwnerDocument.CreateAttribute("DocTypNam");
                            att3.Value = sdr["DocTypNam"].ToString(); 
                            anode.InnerText = _docno + ".TIF";
                            anode.Attributes.Append(att1); anode.Attributes.Append(att2); anode.Attributes.Append(att3);
                            rtnNode.AppendChild(anode);
                        }
                    }
                }
                return rtnNode;
            }
            return null;
        }

        /// <summary>
        /// 借卡及臨時憑證需異動的資料
        /// </summary>
        /// <param name="UserTempCaDetail"></param>
        /// <returns></returns>
        [WebMethod]
        public bool updateTempCaDetail(XmlNode UserTempCaDetail)
        {
            //Key
            string seqno = getUseOrgSeqNO( UserTempCaDetail.SelectSingleNode("//UseOrgSeqNO").InnerText).ToString();
            string docno = UserTempCaDetail.SelectSingleNode("//DocNO").InnerText;
            string flowid = UserTempCaDetail.SelectSingleNode("//FlowId").InnerText;
            //使用臨時憑證時須新增一筆
            string rpsdepid = UserTempCaDetail.SelectSingleNode("//RpsDepID") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsDepID").InnerText;
            string rpsdepnam = UserTempCaDetail.SelectSingleNode("//RpsDepNam") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsDepNam").InnerText;
            string rpsdivid = UserTempCaDetail.SelectSingleNode("//RpsDivID") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsDivID").InnerText;
            string rpsdivnam = UserTempCaDetail.SelectSingleNode("//RpsDivNam") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsDivNam").InnerText;
            string rpsuserid = UserTempCaDetail.SelectSingleNode("//RpsUserID") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsUserID").InnerText;
            string rpsusernam = UserTempCaDetail.SelectSingleNode("//RpsUserNam") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsUserNam").InnerText;
            string rpstime = UserTempCaDetail.SelectSingleNode("//RpsTime") == null ? "" : UserTempCaDetail.SelectSingleNode("//RpsTime").InnerText;
            string applydesc = UserTempCaDetail.SelectSingleNode("//ApplyDesc") == null ? "" : UserTempCaDetail.SelectSingleNode("//ApplyDesc").InnerText;   
            //補簽時需update的資料
            string remdepid = UserTempCaDetail.SelectSingleNode("//RemDepID") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemDepID").InnerText;
            string remdepnam = UserTempCaDetail.SelectSingleNode("//RemDepNam") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemDepNam").InnerText;
            string remdivid = UserTempCaDetail.SelectSingleNode("//RemDivID") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemDivID").InnerText;
            string remdivnam = UserTempCaDetail.SelectSingleNode("//RemDivNam") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemDivNam").InnerText;
            string remuserid = UserTempCaDetail.SelectSingleNode("//RemUserID") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemUserID").InnerText;
            string remusernam = UserTempCaDetail.SelectSingleNode("//RemUserNam") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemUserNam").InnerText;
            string remtime = UserTempCaDetail.SelectSingleNode("//RemTime") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemTime").InnerText;
            string reason = UserTempCaDetail.SelectSingleNode("//RemReason") == null ? "" : UserTempCaDetail.SelectSingleNode("//RemReason").InnerText;
            string cretime = rpstime;
            string rsvA = UserTempCaDetail.SelectSingleNode("//RsvFieldA") == null ? "" : UserTempCaDetail.SelectSingleNode("//RsvFieldA").InnerText;
            string rsvB = UserTempCaDetail.SelectSingleNode("//RsvFieldB") == null ? "" : UserTempCaDetail.SelectSingleNode("//RsvFieldB").InnerText;

            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt),
					new SqlParameter("@DocNO", SqlDbType.Char),
                    new SqlParameter("@FlowId", SqlDbType.NVarChar)};
            Parms[0].Value = seqno;
            Parms[1].Value = docno;
            Parms[2].Value = flowid;
            string sqlstr = "select * from UserTempCaDetail where UseOrgSeqNO = @SeqNO and DocNO = @DocNO and FlowID = @FlowId";

            using (SqlDataReader sdr = SqlHelper.ExecuteReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                //update
                if (sdr.Read())
                {
                    sqlstr = "update UserTempCaDetail set RemDepID = @RemDepID , RemDepNam = @RemDepNam , RemDivID = @RemDivID , RemDivNam = @RemDivNam , RemUserID = @RemUserID , ";
                    sqlstr += "RemUserNam = @RemUserNam , RemTime = @RemTime , RemReason = @RemReason WHERE (UseOrgSeqNO = @SeqNO AND DocNO = @DocNO And FlowID = @FlowId)";
                    SqlParameter[] aParm = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt),
					new SqlParameter("@DocNO", SqlDbType.Char),
                    new SqlParameter("@FlowId", SqlDbType.NChar),
                    new SqlParameter("@RemDepID",SqlDbType.NChar), 
                    new SqlParameter("@RemDepNam",SqlDbType.NChar),
                    new SqlParameter("@RemDivID",SqlDbType.NChar),
                    new SqlParameter("@RemDivNam",SqlDbType.NChar),
                    new SqlParameter("@RemUserID",SqlDbType.NChar),
                    new SqlParameter("@RemUserNam",SqlDbType.NChar),
                    new SqlParameter("@RemTime",SqlDbType.NChar),
                    new SqlParameter("@RemReason",SqlDbType.NChar)};
                    aParm[0].Value = seqno;
                    aParm[1].Value = docno;
                    aParm[2].Value = flowid;
                    aParm[3].Value = remdepid;
                    aParm[4].Value = remdepnam;
                    aParm[5].Value = remdivid;
                    aParm[6].Value = remdivnam;
                    aParm[7].Value = remuserid;
                    aParm[8].Value = remusernam;
                    aParm[9].Value = remtime;
                    aParm[10].Value = reason;
                    int rtn = SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, aParm);
                    if (rtn != 1)
                        return false;
                    else
                        return true;

                }
                else //insert
                {
                    sqlstr = "insert into UserTempCaDetail (UseOrgSeqNO,DocNO,RpsDepID,RpsDepNam,RpsDivID,RpsDivNam,RpsUserID,RpsUserNam,RpsTime,RemDepID,RemDepNam,RemDivID,RemDivNam,RemUserID,RemUserNam,RemTime,FlowID,RemReason,RsvFieldA,RsvFieldB,CreTime,ApplyDesc,ApplyTime,FlowIDDocNo,FlowIDConCur)";
                    sqlstr += " VALUES (@SeqNO,@DocNO,@RpsDepID,@RpsDepNam,@RpsDivID,@RpsDivNam,@RpsUserID,@RpsUserNam,@RpsTime,@RemDepID,@RemDepNam,@RemDivID,@RemDivNam,@RemUserID,@RemUserNam,@RemTime,@FlowId,@RemReason,@RsvFieldA,@RsvFieldB,@CreTime,@applydesc,'','',0)";
                    SqlParameter[] bParm = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt),
					new SqlParameter("@DocNO", SqlDbType.Char),
                    new SqlParameter("@FlowId", SqlDbType.NChar),
                    new SqlParameter("@RemDepID",SqlDbType.NChar), 
                    new SqlParameter("@RemDepNam",SqlDbType.NChar),
                    new SqlParameter("@RemDivID",SqlDbType.NChar),
                    new SqlParameter("@RemDivNam",SqlDbType.NChar),
                    new SqlParameter("@RemUserID",SqlDbType.NChar),
                    new SqlParameter("@RemUserNam",SqlDbType.NChar),
                    new SqlParameter("@RemTime",SqlDbType.NChar),
                    new SqlParameter("@RemReason",SqlDbType.NChar),
                    new SqlParameter("@RpsDepID",SqlDbType.NChar), 
                    new SqlParameter("@RpsDepNam",SqlDbType.NChar),
                    new SqlParameter("@RpsDivID",SqlDbType.NChar),
                    new SqlParameter("@RpsDivNam",SqlDbType.NChar),
                    new SqlParameter("@RpsUserID",SqlDbType.NChar),
                    new SqlParameter("@RpsUserNam",SqlDbType.NChar),
                    new SqlParameter("@RpsTime",SqlDbType.NChar),
                    new SqlParameter("@RsvFieldA",SqlDbType.NChar),
                    new SqlParameter("@RsvFieldB",SqlDbType.NChar),
                    new SqlParameter("@CreTime",SqlDbType.NChar),
                    new SqlParameter("@applydesc",SqlDbType.NVarChar)};
                    bParm[0].Value = seqno;
                    bParm[1].Value = docno;
                    bParm[2].Value = flowid;
                    bParm[3].Value = remdepid;
                    bParm[4].Value = remdepnam;
                    bParm[5].Value = remdivid;
                    bParm[6].Value = remdivnam;
                    bParm[7].Value = remuserid;
                    bParm[8].Value = remusernam;
                    bParm[9].Value = remtime;
                    bParm[10].Value = reason;
                    bParm[11].Value = rpsdepid;
                    bParm[12].Value = rpsdepnam;
                    bParm[13].Value = rpsdivid;
                    bParm[14].Value = rpsdivnam;
                    bParm[15].Value = rpsuserid;
                    bParm[16].Value = rpsusernam;
                    bParm[17].Value = rpstime;
                    bParm[18].Value = "";
                    bParm[19].Value = "";
                    bParm[20].Value = cretime;
                    bParm[21].Value = applydesc;
                         int rtn = SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, bParm);
                    if (rtn != 1)
                        return false;
                    else
                        return true;
                }
            }
        }

        /// <summary>
        /// 取得彙併辦公文文號
        /// </summary>
        /// <param name="_useorgseqno">機關代碼</param>
        /// <param name="_docno">文號</param>
        /// <returns>第一筆為母文其餘子文</returns>
        [WebMethod]
        public List<string> getMergeData(string _orgid, string _docno)
        {
            string _useorgseqno = Convert.ToString(getUseOrgSeqNO(_orgid));
            string sqlstr = "select * from CombineDetail where ParDocNO = @DocNO AND UseOrgSeqNO = @SeqNO";
            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@SeqNO", SqlDbType.SmallInt,2),
					new SqlParameter("@DocNO", SqlDbType.Char, _docno.Length)};
            Parms[0].Value = _useorgseqno;
            Parms[1].Value = _docno;
            List<string> rtn = new List<string>();
            using (SqlDataReader sdr = SqlHelper.ExecuteReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
            {
                if (sdr.HasRows)
                {
                    rtn.Add(_docno);
                    while (sdr.Read())
                    {
                        rtn.Add(sdr["ChiDocNO"].ToString());
                    }
                }
                else
                {
                    sqlstr = "select * from CombineDetail where ChiDocNO = @DocNO AND UseOrgSeqNO = @SeqNO";
                    using (SqlDataReader dr = SqlHelper.ExecuteReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms))
                    {
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                if (!rtn.Contains(dr["ParDocNO"].ToString()))
                                    rtn.Add(dr["ParDocNO"].ToString());

                                if (!rtn.Contains(dr["ChiDocNO"].ToString()))
                                    rtn.Add(dr["ChiDocNO"].ToString());
                            }
                        }
                    }
                }
            }
            return rtn;
        }

        /// <summary>
        /// 取得線上簽核si檔內容
        /// </summary>
        /// <param name="_docno">文號</param>
        /// <returns>null表示尚無裝檔,用來判斷是否須新建</returns>
        [WebMethod]
        public XmlDocument GetSiXML(string _docno,string orgid)
        {
            return readSiFile(_docno,orgid);
        }

        [WebMethod]
        public bool SiExist(string _docno, string orgid)
        {
            //string sipath = Path.Combine(ConfigurationSettings.AppSettings["siRootPath"], _docno.Substring(0, 3));
            //sipath = Path.Combine(sipath, _docno.Substring(_docno.Length - 2, 2));
            ////sipath = Path.Combine(sipath, _docno.Substring(5));
            //sipath = Path.Combine(sipath, _docno);
            string sipath = getsipath(_docno, orgid);
            if (_DepId != "")
                sipath = Path.Combine(sipath, _DepId);
            SiFilePath = sipath;
            if (!File.Exists(Path.Combine(SiFilePath, _docno + ".dsicdp")))
                return false;
            else
                return true;
        }

        /// <summary>
        /// 取出並會的si
        /// </summary>
        /// <param name="_docno"></param>
        /// <param name="depid"></param>
        /// <returns></returns>
        [WebMethod]
        public XmlDocument GetDepSiXML(string _docno,string depid,string orgid)
        {
            _DepId = depid;
            return readSiFile(_docno,orgid);
        }

        /// <summary>
        /// 將補簽後的的si檔暫存在Resign資料夾
        /// </summary>
        /// <param name="_docno">文號</param>
        /// <param name="resignID">補簽ID</param>
        /// <param name="siDoc">簽完後的xml</param>
        /// <returns></returns>
        [WebMethod]
        public bool SaveReSignXML(string _docno,string orgid, string resignID, XmlDocument siDoc)
        {
            //擺放路徑 前3碼 + 4,5碼 + 餘碼 + 文號
            //string sipath = Path.Combine(ConfigurationSettings.AppSettings["siRootPath"], _docno.Substring(0, 3));
            //if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            //sipath = Path.Combine(sipath, _docno.Substring(_docno.Length -2 , 2));
            //if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            ////sipath = Path.Combine(sipath, _docno.Substring(5));
            ////if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            //sipath = Path.Combine(sipath, _docno);
            //if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            string sipath = getsipath(_docno, orgid);
            sipath = Path.Combine(sipath, "ReSign");
            if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            sipath = Path.Combine(sipath, resignID);
            if (File.Exists(sipath)) File.Delete(sipath);
            siDoc.Save(sipath);
            return true;
        }

        /// <summary>
        /// 將該文所有補簽資料並回主si檔
        /// </summary>
        /// <param name="_docno">文號</param>
        /// <returns></returns>
        [WebMethod]
        public bool MergeReSignXML(string _docno,string orgid)
        {
            XmlDocument orgXml = readSiFile(_docno,orgid);
            if (orgXml == null) return false;
            XmlNodeList resignNode = orgXml.DocumentElement.GetElementsByTagName("補簽追認");
            //先將現有的節點刪除重新產生
            if (resignNode.Count == 1) orgXml.DocumentElement.RemoveChild(resignNode[0]);

            //取出補的si檔並且併入原si檔
            DirectoryInfo dir = new DirectoryInfo(Path.Combine(SiFilePath, "ReSign"));
            try
            {
                if (dir.GetFiles().Length > 0)
                {
                    XmlNode _補簽追認 = orgXml.CreateElement("補簽追認") as XmlNode;

                    int recnt = 0;
                    foreach (FileInfo f in dir.GetFiles())
                    {
                        XmlDocument reXml = new XmlDocument();
                        reXml.XmlResolver = null;
                        reXml.Load(f.FullName);
                        reXml.PreserveWhitespace = true;
                        foreach (XmlNode x in reXml.DocumentElement.GetElementsByTagName("補簽人員"))
                        {
                            _補簽追認.AppendChild(_補簽追認.OwnerDocument.ImportNode(x, true));
                            recnt++;
                        }
                    }
                    XmlAttribute _補簽人員數 = orgXml.CreateAttribute("補簽人員數");
                    _補簽人員數.Value = recnt.ToString();
                    _補簽追認.Attributes.Append(_補簽人員數);
                    orgXml.DocumentElement.GetElementsByTagName("線上簽核資訊")[0].AppendChild(orgXml.ImportNode(_補簽追認, true));
                    orgXml.Save(Path.Combine(SiFilePath, "unzip") + "\\" + _docno + ".si");
                    //將改完的資料再壓縮回檔案
                    DCPWebService.GZip.Compress(Path.Combine(SiFilePath, "unzip"), SiFilePath, _docno + ".dsicdp");
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 使用HSM做加簽
        /// </summary>
        /// <param name="digestXml">已算過HASH的XML</param>
        /// <param name="sigid">計算加簽的tag id</param>
        /// <param name="keyname">HSM裡的機關憑證id</param>
        /// <returns>加簽過的xml字串</returns>
        [WebMethod]
        public string SignWithHSM(string digestXml, string sigid, string keyname)
        {
            string servletURL = ConfigurationSettings.AppSettings["HSMFullPath"].ToString();//@"http://172.16.1.133:8080/SSServlet/SS";
            string apiUser = "api";
            string apiPass = "A033A528B603FED46F861D4B3542C417B99D41C8";

            RA2AGENTATLLib.RA2FacadeClass ss = new RA2AGENTATLLib.RA2FacadeClass();

            int rtn = ss.FSRA2_ServerAddURL(servletURL);
            if (rtn != 0)
            {
                string errMsg = ss.FSRA2_GetErrorMsg();
                return "err1:" + errMsg ;
            }

            rtn = ss.FSRA2_Login(apiUser, apiPass, "This is a log.");
            if (rtn != 0)
            {
                string errMsg = ss.FSRA2_GetErrorMsg();
                return "err2:" + errMsg;
            }

            string key1 = ss.FSRA2_GetKey1();
            string keyName = keyname != "" ? keyname : "mof";
            string apName = "DsicOP";
            int FS_XML_FLAG_IGNORE_DIGESTVALUE = 0x08000000;
            int iFlag = FS_XML_FLAG_IGNORE_DIGESTVALUE;
            //int iFlag = 0;
            
            rtn = ss.FSSS_XMLSignByTemplateDigest(key1, keyName, digestXml, sigid, apName, iFlag, "This is a log.");
            if (rtn != 0)
            {
                string errMsg = ss.FSRA2_GetErrorMsg();
                return "err3:" + errMsg;
            }
            string signxml = ss.FSSS_GetSignature();
            return signxml;
        }

        /// <summary>
        /// 註記封裝是否有完成
        /// </summary>
        /// <param name="modecode">模式代碼</param>
        /// <param name="flowid">流程點ID</param>
        /// <returns></returns>
        [WebMethod]
        public bool UpdateStatus(string modecode, string flowid, string oedocno, string userid, string _orgid, string DocNO)
        {
            int useseqno = getUseOrgSeqNO(_orgid);
            string _useorgseqno = Convert.ToString(getUseOrgSeqNO(_orgid));
            string tblname = "";
            string wherestr = "";
            switch (modecode)
            {
                case "1":
                case "2":
                case "3":
                case "mergeDep":
                    tblname = "RpsDepDetail";
                    wherestr = "and ModTime = (Select Max(ModTime) from RpsDepDetail where UseOrgSeqNO = " + _useorgseqno + " and OEDocNO='" + oedocno + "' and RpsUserID='" + userid + "')";
                    break;
                case "6":
                case "B":
                    tblname = "DepRegDetail";
                    wherestr = "and ModTime = (Select Max(ModTime) from DepRegDetail where UseOrgSeqNO = " + _useorgseqno + " and OEDocNO='" + oedocno + "' and RpsUserID='" + userid + "')";
                    break;
                //case "B":
                //    tblname = "ReviewDetail";
                //    wherestr = "and ModTime = (Select Max(ModTime) from ReviewDetail where OEDocNO='" + oedocno + "')";
                //    break;
                case "7":
                case "8":
                case "G":
                    tblname = "JudgeDetail";
                    wherestr = "and JudTime = (Select Max(JudTime) from JudgeDetail where UseOrgSeqNO = " + _useorgseqno + " and OEDocNO='" + oedocno + "' and JudUserID='" + userid + "')";
                    break;
                case "C":
                case "9":
                    return true;
                    break;
                default:
                    break;
            }
            string sqlstr = "update " + tblname + " set flowid = @flowid where OEDocNO = @oedocno and UseOrgSeqNO = @useorgno " + wherestr;
            SqlParameter[] Parms = new SqlParameter[] {
                    new SqlParameter("@oedocno", SqlDbType.Char),
                    new SqlParameter("@flowid", SqlDbType.NChar),
                    new SqlParameter("@useorgno", SqlDbType.Int)};
            Parms[0].Value = oedocno;
            Parms[1].Value = flowid;
            Parms[2].Value = useseqno;
            int cnt = SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms);
            if (cnt == 1)
            {
                if (modecode == "mergeDep")
                {
                    sqlstr = "update RegDepDivsDetail set RegCtrl='1' where DocNO = @docno and UseOrgSeqNO =  @useorgno";
                    Parms = new SqlParameter[]{
                        new SqlParameter("@docno", SqlDbType.Char),
                        new SqlParameter("@useorgno", SqlDbType.Int)};

                    Parms[0].Value = DocNO;
                    Parms[1].Value = useseqno;
                    int rnt = SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sqlstr, Parms);
                    if (cnt == 1)
                    {
                        return true;
                    }
                    else
                        return false;
                }
                else
                    return true;
            }
            else
                return false;
        }

        /// <summary>
        /// 讀取簽核電子檔XML內容
        /// </summary>
        /// <param name="_docno">公文文號</param>
        /// <returns></returns>
        private XmlDocument readSiFile(string _docno, string orgid)
        {
            //string sipath = Path.Combine(ConfigurationSettings.AppSettings["siRootPath"], _docno.Substring(0, 3));
            //sipath = Path.Combine(sipath, _docno.Substring(_docno.Length - 2, 2));
            ////sipath = Path.Combine(sipath, _docno.Substring(5));
            //sipath = Path.Combine(sipath, _docno);
            string sipath = getsipath(_docno, orgid);
            if (_DepId != "")
                sipath = Path.Combine(sipath, _DepId);
            SiFilePath = sipath;
            if (!File.Exists(Path.Combine(SiFilePath, _docno + ".dsicdp"))) return null;

            GZipResult gr = DCPWebService.GZip.Decompress(sipath, Path.Combine(sipath, "unzip"), _docno + ".dsicdp");
            gr = null;
            if (!File.Exists(Path.Combine(sipath, "unzip") + "\\" + _docno + ".si"))
            {
                return null;
            }
            //取得原本的si檔
            XmlDocument orgXml = new XmlDocument();
            orgXml.XmlResolver = null;
            orgXml.Load(Path.Combine(sipath, "unzip") + "\\" + _docno + ".si");
            //orgXml.InsertBefore(orgXml.CreateXmlDeclaration("1.0", "utf-8", null), orgXml.FirstChild);
            orgXml.PreserveWhitespace = true;
            //DeletePathFiles(Path.Combine(sipath, "unzip"));
            return orgXml;
        }

        /// <summary>
        /// 依文號取得封裝檔路徑
        /// </summary>
        /// <param name="_docno">公文文號</param>
        /// <returns></returns>
        private string getsipath(string _docno,string orgid)
        {
            //擺放路徑 前3碼 + 後2 + 文號
            string sipath = Path.Combine(ConfigurationSettings.AppSettings["siRootPath"], orgid);
            if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            sipath = Path.Combine(sipath, _docno.Substring(0, 3));
            if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            sipath = Path.Combine(sipath, _docno.Substring(_docno.Length-2 , 2));
            if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            //sipath = Path.Combine(sipath, _docno.Substring(5));
            //if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            sipath = Path.Combine(sipath, _docno);
            if (!Directory.Exists(sipath)) Directory.CreateDirectory(sipath);
            return sipath;
        }

        #region 檔案傳輸功能
        /// <summary>
        /// 上傳封裝的壓縮檔
        /// </summary>
        /// <param name="_docno">文號</param>
        /// <param name="buffer">檔案內容</param>
        /// <param name="Offset"></param>
        [WebMethod]
        public void UploadChunk(string _docno,  byte[] buffer, long Offset,string orgid)
        {
            string FileName = _docno + ".dsicdp";
            string FilePath = System.IO.Path.Combine(getsipath(_docno,orgid), FileName);

            if (Offset == 0)	// new file, create an empty file
                System.IO.File.Create(FilePath).Close();

            // open a file stream and write the buffer.  Don't open with FileMode.Append because the transfer may wish to start a different point
            using (System.IO.FileStream fs = new System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.Read))
            {
                fs.Seek(Offset, System.IO.SeekOrigin.Begin);
                fs.Write(buffer, 0, buffer.Length);
            }
        }

        [WebMethod]
        public void UploadDepChunk(string _docno, byte[] buffer, long Offset, string foldername, string orgid)
        {
            string FileName = _docno + ".dsicdp";
            string target = System.IO.Path.Combine(getsipath(_docno,orgid), foldername);
            CreatePath(target);
            string FilePath = System.IO.Path.Combine(target, FileName);

            if (Offset == 0)	// new file, create an empty file
                System.IO.File.Create(FilePath).Close();

            // open a file stream and write the buffer.  Don't open with FileMode.Append because the transfer may wish to start a different point
            using (System.IO.FileStream fs = new System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.Read))
            {
                fs.Seek(Offset, System.IO.SeekOrigin.Begin);
                fs.Write(buffer, 0, buffer.Length);
            }
        }
        
        /// <summary>
        /// 下載壓縮的封裝檔
        /// </summary>
        /// <param name="_docno"></param>
        /// <param name="Offset"></param>
        /// <param name="BufferSize"></param>
        /// <returns></returns>
        [WebMethod]
        public byte[] DownloadChunk(string _docno, string orgid)
        {
            long Offset = 0;
            int BufferSize;
            string FileName = _docno + ".dsicdp";
            string FilePath = System.IO.Path.Combine(getsipath(_docno,orgid), FileName);

            long FileSize = new System.IO.FileInfo(FilePath).Length;
            BufferSize = (int)FileSize;
            byte[] TmpBuffer;
            int BytesRead;

            try
            {
                using (System.IO.FileStream fs = new System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read))
                {
                    fs.Seek(Offset, System.IO.SeekOrigin.Begin);	// this is relevent during a retry. otherwise, it just seeks to the start
                    TmpBuffer = new byte[BufferSize];
                    BytesRead = fs.Read(TmpBuffer, 0, BufferSize);	// read the first chunk in the buffer (which is re-used for every chunk)
                }
                if (BytesRead != BufferSize)
                {
                    // the last chunk will almost certainly not fill the buffer, so it must be trimmed before returning
                    byte[] TrimmedBuffer = new byte[BytesRead];
                    Array.Copy(TmpBuffer, TrimmedBuffer, BytesRead);
                    return TrimmedBuffer;
                }
                else
                    return TmpBuffer;
            }
            catch
            {
                return null;
            }
        }

        [WebMethod]
        public byte[] DownloadDepChunk(string _docno, string foldername, string orgid)
        {
            
            string FileName = _docno + ".dsicdp";
            //第一個會辦人需重承辦處讀取檔案
            if (!Directory.Exists(System.IO.Path.Combine(getsipath(_docno,orgid), foldername)))
            {
                return DownloadChunk(_docno,orgid);
            }
            string FilePath = System.IO.Path.Combine(System.IO.Path.Combine(getsipath(_docno,orgid), foldername), FileName);
            long Offset = 0;
            int BufferSize;
            // check that requested file exists
            if (!System.IO.File.Exists(FilePath))
                return null;

            long FileSize = new System.IO.FileInfo(FilePath).Length;
            BufferSize = (int)FileSize;
            // if the requested Offset is larger than the file, quit.
            if (Offset > FileSize)
                return null;

            // open the file to return the requested chunk as a byte[]
            byte[] TmpBuffer;
            int BytesRead;

            try
            {
                using (System.IO.FileStream fs = new System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read))
                {
                    fs.Seek(Offset, System.IO.SeekOrigin.Begin);	// this is relevent during a retry. otherwise, it just seeks to the start
                    TmpBuffer = new byte[BufferSize];
                    BytesRead = fs.Read(TmpBuffer, 0, BufferSize);	// read the first chunk in the buffer (which is re-used for every chunk)
                }
                if (BytesRead != BufferSize)
                {
                    // the last chunk will almost certainly not fill the buffer, so it must be trimmed before returning
                    byte[] TrimmedBuffer = new byte[BytesRead];
                    Array.Copy(TmpBuffer, TrimmedBuffer, BytesRead);
                    return TrimmedBuffer;
                }
                else
                    return TmpBuffer;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 取得server上檔案大小
        /// </summary>
        /// <param name="_docno"></param>
        /// <returns></returns>
        [WebMethod]
        public long GetFileSize(string _docno, string orgid)
        {
            string FileName = _docno + ".dsicdp";
            string FilePath = System.IO.Path.Combine(getsipath(_docno,orgid), FileName);

            // check that requested file exists
            if (!System.IO.File.Exists(FilePath))
                return -1;

            return new System.IO.FileInfo(FilePath).Length;
        }

        [WebMethod]
        public long GetDepFileSize(string _docno, string DepId, string orgid)
        {
            string FileName = _docno + ".dsicdp";
            if (!Directory.Exists(Path.Combine(getsipath(_docno,orgid), DepId)))
            {
                return GetFileSize(_docno,orgid);
            }
            string FilePath = System.IO.Path.Combine(Path.Combine(getsipath(_docno,orgid), DepId), FileName);

            // check that requested file exists
            if (!System.IO.File.Exists(FilePath))
                return -1;

            return new System.IO.FileInfo(FilePath).Length;
        }

        [WebMethod]
        public List<string> GetFilesList(string UploadPath, out bool DirectoryExist)
        {
            System.IO.DirectoryInfo df;
            List<string> files = new List<string>();

            DirectoryExist = true;
            try
            {
                df = new System.IO.DirectoryInfo(UploadPath);
                if (!df.Exists)
                {
                    DirectoryExist = false;
                    goto Exit;
                }

                foreach (string s in System.IO.Directory.GetFiles(UploadPath))
                {
                    try
                    {
                        files.Add(System.IO.Path.GetFileName(s));
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch
            {
                try
                {
                    DirectoryExist = false;
                    files.Clear();
                }
                catch
                {
                    files = null;
                }
                goto Exit;
            }

        Exit:
            return files;
        }

        /// <summary>
        /// 並會整併後刪除所留下的單位資料夾
        /// </summary>
        /// <param name="_docno"></param>
        /// <param name="DepId"></param>
        /// <param name="orgid"></param>
        /// <returns></returns>
        [WebMethod]
        public string DeleteDepFolder(string _docno, string DepId, string orgid)
        {
            try
            {
                string[] dep = DepId.Split(new char[] { ',' });
                for (int i = 0; i < dep.Length; i++)
                {
                    string target = Path.Combine(getsipath(_docno, orgid), dep[i]);
                    if (Directory.Exists(target))
                    {
                        string a = DeletePathFiles(target);
                        if (a != "") return a;
                        if (!DeletePath(target)) return "刪除" + target + "資料夾失敗!"; ;
                    }
                }
                return "";
            }
            catch (Exception err)
            {
                return err.Message;
            }
        }

        private bool CreatePath(string ServerPath)
        {
            bool bRes = true;
            System.IO.DirectoryInfo df;

            try
            {
                df = new System.IO.DirectoryInfo(ServerPath);
                if (!df.Exists)
                {
                    df = System.IO.Directory.CreateDirectory(ServerPath);
                }
            }
            catch
            {
                bRes = false;
                goto Exit;
            }

        Exit:
            return bRes;
        }

        private bool DeletePath(string ServerPath)
        {
            bool bRes = true;
            System.IO.DirectoryInfo df;

            try
            {
                df = new System.IO.DirectoryInfo(ServerPath);
                if (df.Exists)
                {
                    System.IO.Directory.Delete(ServerPath, true);
                }
            }
            catch
            {
                bRes = false;
                goto Exit;
            }

        Exit:
            return bRes;
        }

        private string DeletePathFiles(string ServerPath)
        {
            bool bRes = true;
            System.IO.DirectoryInfo df;

            try
            {
                df = new System.IO.DirectoryInfo(ServerPath);
                if (df.Exists)
                {
                    foreach (string s in System.IO.Directory.GetFiles(ServerPath))
                    {
                        try
                        {
                            System.IO.File.Delete(s);
                        }
                        catch (Exception err)
                        {
                            return err.Message;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                bRes = false;
                return err.Message;
                goto Exit;
            }

        Exit:
            return "";
            //return bRes;
        }

        #endregion

        #region 行動簽核
        [WebMethod]
        public XmlNode getXmlforSign(string DocNO,string OEDocNO,string OrgID,string depnam, string depid, string usertitle, string usernam, string userid, string comment, string judgestr)
        {
            string wseurl = this.Context.Request.Url.AbsoluteUri;
            string dsicdppath =Path.Combine(getsipath(DocNO,OrgID),DocNO + ".dsicdp");
            MobileDCP mdcp = new MobileDCP(wseurl, dsicdppath, OrgID, DocNO, OEDocNO);
            string rtn = mdcp.MakeDCP(depnam, depid, usertitle, usernam, userid, comment, judgestr);
            XmlNode xnd = new XmlDocument().CreateElement("SignXml");

            if (!string.IsNullOrEmpty(rtn))
            {
                xnd.InnerText = "Error:呼叫MakeDCP發生錯誤，錯誤訊息為" + rtn;
            }
            else
            {
                XmlAttribute flowid = xnd.OwnerDocument.CreateAttribute("flowid");
                flowid.Value = "sign_" + mdcp.FlowID;
                xnd.Attributes.Append(flowid);
                XmlAttribute signid = xnd.OwnerDocument.CreateAttribute("signid");
                signid.Value = mdcp.CurrentSignId;
                xnd.Attributes.Append(signid);
                XmlAttribute modeid = xnd.OwnerDocument.CreateAttribute("modeid");
                modeid.Value = string.IsNullOrEmpty(judgestr) ? "7" : "8";
                xnd.Attributes.Append(modeid);
                xnd.InnerText = mdcp.ForSignXML;
            }
            return xnd;
        }

        [WebMethod]
        public string updateDCPData(string OrgID, string DocNO,string oedocno, string signature, string cert, string signid,string flowid,string userid,string modeid)
        {
            //將簽體憑證更新回si檔
            XmlDocument xdom = new XmlDocument();
            string sipath = getsipath(DocNO, OrgID) + "\\DcpSource\\" + DocNO + ".si";
            string tmppath = getsipath(DocNO, OrgID) + "\\DcpTemp\\";
            xdom.XmlResolver = null;
            xdom.Load(sipath);
            xdom.PreserveWhitespace = true;
            DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature sig = new DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature();
            string bstr = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(xdom.OuterXml));
            string rtnstr = sig.UpdateSignatureValue(bstr, signature, cert, signid);
            if (rtnstr.StartsWith("ErrA:"))
                return "Error:UpdateSignatureValue失敗，錯誤訊息為" + rtnstr;
            xdom.LoadXml(rtnstr);
            xdom.Save(sipath);
            //將封裝資料壓縮並放回原位置
            MakeDCP mkDCP = new MakeDCP(DocNO);
            string rtn = mkDCP.壓縮封裝檔(sipath, tmppath);
            if (string.IsNullOrEmpty(rtn))
            {
                File.Copy(Path.Combine(tmppath, DocNO + ".dsicdp"), Path.Combine(getsipath(DocNO, OrgID),DocNO + ".dsicdp"), true);
                if (!UpdateStatus(modeid, flowid, oedocno, userid, OrgID, DocNO))
                {
                    return "Error:DB更新失敗。";
                }
                return "";
            }
            else
                return "Error:壓縮失敗，" + rtn;
        }
        #endregion
    }
}