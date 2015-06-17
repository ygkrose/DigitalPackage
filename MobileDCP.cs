using System;
using System.Collections.Generic;
using System.Text;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo;
using DigitalSealed.Tools;
using System.IO;
using DSIC.WEBOE.MAIN.Controller;
using System.Xml;

namespace DigitalSealed
{
    public class MobileDCP
    {
        private string dcpXML = "";
        /// <summary>
        /// 完整DCP XML
        /// </summary>
        public string DCPXML
        {
            get { return dcpXML; }
        }

        private string currentid = "";
        /// <summary>
        /// 待加簽簽核點ID
        /// </summary>
        public string CurrentSignId
        {
            get { return currentid; }
        }

        private string forsigxml = "";
        /// <summary>
        /// 待加簽的XML
        /// </summary>
        public string ForSignXML
        {
            get { return forsigxml; }
        }

        private string flowid = "";
        /// <summary>
        /// 簽核點ID
        /// </summary>
        public string FlowID
        {
            get { return flowid; }
            set { flowid = value; }
        }

        private ModifyInfo MInfo = null;
        private ModifyInfo.signer sr = null;
        private string _dsicdp = "";
        private string sourcefilepath = "";
        private string dcptemppath = "";
        private bool transimg = false;
        private string ExeTyp = "";
        private MakeDCP mdcp = null;
        private FileInfo dcpfile = null;
        private string OrgID = "";
        private string DocNO = "";
        private string OEDocNO = "";
        private string DCPFileTransferWebServiceURL = "";

        /// <summary>
        /// mobileDCP建構子
        /// </summary>
        /// <param name="DcpFileWSE">DCP WSE URL</param>
        /// <param name="FullDsicdpPath">DSICDP封裝檔位置</param>
        /// <param name="orgid">機關代碼</param>
        /// <param name="docno">文號</param>
        /// <param name="oedocno">文書號</param>
        public MobileDCP(string DcpFileWSE,string FullDsicdpPath,string orgid,string docno,string oedocno) 
        {
            if (File.Exists(FullDsicdpPath))
            {
                try
                {
                    OrgID = orgid;
                    DocNO = docno;
                    OEDocNO = oedocno;
                    _dsicdp = FullDsicdpPath;
                    DCPFileTransferWebServiceURL = DcpFileWSE;
                    dcpfile = new FileInfo(_dsicdp);
                    Directory.CreateDirectory(Path.Combine(dcpfile.DirectoryName, "DcpSource"));
                    Directory.CreateDirectory(Path.Combine(dcpfile.DirectoryName, "DcpTemp"));
                    sourcefilepath = Path.Combine(dcpfile.DirectoryName, "DcpSource");
                    dcptemppath = Path.Combine(dcpfile.DirectoryName, "DcpTemp");
                }
                catch (Exception err)
                {
                    throw new Exception("建立資料夾失敗" + err.Message);
                }
            }
            else
                throw new Exception("封裝檔" + FullDsicdpPath + "不存在");
        }

        /// <summary>
        /// 取得算完HASH待加簽的XML字串
        /// </summary>
        /// <param name="depnam"></param>
        /// <param name="depid"></param>
        /// <param name="usertitle"></param>
        /// <param name="usernam"></param>
        /// <param name="userid"></param>
        /// <param name="comment"></param>
        /// <param name="judgestr"></param>
        /// <returns></returns>
        public string MakeDCP(string depnam, string depid, string usertitle, string usernam, string userid, string comment, string judgestr)
        {
            string rtn = "";
            sr = new DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo.ModifyInfo.signer(depnam, depid, usertitle, usernam, userid);
            if (judgestr != "")
            {
                MInfo = new ModifyInfo(sr, FlowInfo.FlowType.決行, comment, geneTime.getTimeNow());
                transimg = true;
                ExeTyp = "8";
                rtn = basicproc(FlowInfo.FlowType.決行, comment);
            }
            else
            {
                MInfo = new ModifyInfo(sr, FlowInfo.FlowType.批示, comment, geneTime.getTimeNow());
                ExeTyp = "7";
                rtn = basicproc(FlowInfo.FlowType.批示, comment);
            }
            
            
            if (string.IsNullOrEmpty(rtn))
            {
                SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature sig = new DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature();
                dcpXML = mdcp.封裝檔XML.OuterXml;
                flowid = mdcp.簽核點ID;
                mdcp.封裝檔XML.Save(Path.Combine(sourcefilepath, DocNO + ".si"));
                //mdcp.封裝檔XML.Save(@"d:\temp\orgsi.txt");
                currentid = "ds_" + mdcp.簽核點ID;
                try
                {
                    //using (TextWriter writer = new StreamWriter(@"d:\temp\bastrsi.txt"))
                    //{
                    //    writer.Write(Convert.ToBase64String(Encoding.UTF8.GetBytes(mdcp.封裝檔XML.OuterXml)));
                    //}
                    forsigxml = sig.GetSignedInfoByTemplateDigest(Convert.ToBase64String(Encoding.UTF8.GetBytes(mdcp.封裝檔XML.OuterXml)), currentid);
                    forsigxml = Convert.ToBase64String(Encoding.UTF8.GetBytes(forsigxml));
                    //using (TextWriter writer = new StreamWriter(@"d:\temp\c14nsi.txt"))
                    //{
                    //    writer.Write(forsigxml);
                    //}
                }
                catch (Exception err)
                {
                    rtn = err.Message;
                }
                
                //XmlNode elm = mdcp.封裝檔XML.SelectSingleNode("//Signature[@Id='" + "ds_" + mdcp.簽核點ID + "']");
                //if (elm != null)
                //{
                //    XmlNode sinfo = elm.SelectSingleNode("//SignedInfo");
                //    if (sinfo != null)
                //    {
                //        forsigxml = Convert.ToBase64String(Encoding.UTF8.GetBytes(sinfo.OuterXml));
                //    }
                //}
            }
            return rtn;
        }


        private string fileCompress(MakeDCP mdcp)
        {
            //將封裝檔存檔並刪除暫存檔
            mdcp.封裝檔XML.Save(Path.Combine(mdcp.檔案來源資料夾, mdcp.公文文號 + ".si"));

            //將所有封裝內容做壓縮
            string rtn = mdcp.壓縮封裝檔(Path.Combine(mdcp.檔案來源資料夾, mdcp.公文文號 + ".si"), dcptemppath);
            return rtn;
        }

        private bool updateFlowIDStatus()
        {
            object[] webMethodparam = new object[6];
            webMethodparam[0] = ExeTyp;
            webMethodparam[1] = mdcp.簽核點ID;
            webMethodparam[2] = OEDocNO;
            webMethodparam[3] = sr.帳號;
            webMethodparam[4] = OrgID;
            webMethodparam[5] = DocNO;
            //string exetype = _IntegrationInfo.sysConfig.ExeTyp;
            //return wse.UpdateStatus(exetype, mdcp.簽核點ID, OEDocNO, _IntegrationInfo.sysConfig.UserID, OrgID, DocNO);
            return (bool)WebServiceController.InvokeWebservice(DCPFileTransferWebServiceURL, "DCPWebService", "DCPWSE", "UpdateStatus", webMethodparam);
        }

        #region 主程序
        /// <summary>
        /// 傳遞程序
        /// </summary>
        /// <returns></returns>
        private string basicproc(FlowInfo.FlowType modename,string Desc)
        {
            mdcp = new MakeDCP(dcpfile,new DirectoryInfo(sourcefilepath), sr, null, Desc.Trim().Replace("&nbsp;", ""), modename);
            mdcp.暫存資料夾 = dcptemppath;
            mdcp.檔案來源資料夾 = sourcefilepath;
            //mdcp.HashAlgorithm = "SHA256";
            //將文面與簽核意見存到來源資料夾
            //if (File.Exists(Path.Combine(CommonTool.CheckoutfilePath, OEDocNO + "withSignInfo.tif")))
            //{
            //    File.Copy(Path.Combine(CommonTool.CheckoutfilePath, OEDocNO + "withSignInfo.tif"), Path.Combine(sourcefilepath, OEDocNO + ".tif"), true);
            //}
            //首先處理來文
            string result = "";

            result = mdcp.產生來文參考();

            //處理文書及附件
            result += geneOEDoc(modename.ToString());

             //if (modename != FlowInfo.FlowType.解併 && modename != FlowInfo.FlowType.彙併辦 && _IntegrationInfo.相關資訊.使用者資訊.借卡時間 != "")
            //{
            //    if (!updateTempCardStatus())
            //        result = "更新待補簽資料庫失敗!";
            //}
            if (result == "")
                result = mdcp.異動封裝檔("", transimg); //密碼帶空值只算雜湊

            if (result == "")
            {
                DeleteUnusedFiles();
            }
            return result;
        }

        List<string> multiOEDocNO = null;
        private string geneOEDoc(string modeType)
        {
            string rtn = "";
            object[] webMethodparam = new object[2];
            webMethodparam[0] = OrgID;
            webMethodparam[1] = DocNO;
            //XmlNode rtnNode = wse.getWordDocAttInfo(OrgID, DocNO);
            XmlNode rtnNode = WebServiceController.InvokeWebservice(DCPFileTransferWebServiceURL, "DCPWebService", "DCPWSE", "getWordDocAttInfo", webMethodparam) as XmlNode;
            multiOEDocNO = new List<string>();
            bool isjudge = false;
            foreach (XmlNode xnd in rtnNode.ChildNodes)
            {
                multiOEDocNO.Add(xnd.Attributes["OEDocNO"].Value.Trim());
                //if (xnd.Attributes[0].Value.Trim() == OEDocNO)
                //{
                //    string difile = Path.Combine(Path.Combine(CommonTool.CheckoutfilePath, OEDocNO), OEDocNO + ".di");
                //    string difilever = OEDocNO + "V" + xnd.Attributes[2].Value.Trim() + ".di";
                //    File.Copy(difile, Path.Combine(sourcefilepath, difilever), true);
                //    //處理附件
                //    List<string> atts = new List<string>();
                //    foreach (XmlNode chnd in xnd.ChildNodes)
                //    {
                //        if (chnd.Attributes.Count == 0) continue;
                //        if (File.Exists(Path.Combine(Path.Combine(CommonTool.CheckoutfilePath, OEDocNO), chnd.Attributes["AtaFile"].Value.Trim())))
                //        {
                //            FileInfo sourceAtt = new FileInfo(Path.Combine(Path.Combine(CommonTool.CheckoutfilePath, OEDocNO), chnd.Attributes["AtaFile"].Value.Trim()));
                //            string targetAtt = Path.Combine(sourcefilepath, sourceAtt.Name.Replace(sourceAtt.Extension, "") + "V" + chnd.Attributes["LastChangeVer"].Value.Trim() + sourceAtt.Extension);
                //            string targetFileName = sourceAtt.Name.Replace(sourceAtt.Extension, "") + "V" + chnd.Attributes["LastChangeVer"].Value.Trim() + sourceAtt.Extension;
                //            //string outstr = chkSameAttName(chnd.Attributes["AtaFile"].Value.Trim(), sourceAtt);
                //            //if (outstr != "") return outstr;
                //            File.Copy(sourceAtt.FullName, targetAtt, true);
                //            atts.Add(targetFileName);
                //        }
                //        else
                //        {
                //            string oname = chnd.Attributes["AtaFile"].Value.Trim();
                //            string oextname = oname.Substring(oname.LastIndexOf("."));
                //            string fname = oname.Replace(oextname, "") + "V" + chnd.Attributes["LastChangeVer"].Value.Trim() + oextname;
                //            if (File.Exists(Path.Combine(sourcefilepath, fname)))
                //                atts.Add(fname);
                //            else
                //            {
                //                FileInfo sourceAtt = downloadOEatt(this.dcptemppath, oname, chnd.Attributes["LastChangeVer"].Value.Trim());
                //                if (sourceAtt == null) return "下載文書附件失敗!";
                //                string targetAtt = Path.Combine(sourcefilepath, sourceAtt.Name.Replace(sourceAtt.Extension, "") + "V" + chnd.Attributes["LastChangeVer"].Value.Trim() + sourceAtt.Extension);
                //                File.Copy(this.dcptemppath + oname, targetAtt);
                //                atts.Add(fname);
                //            }
                //        }

                //    }
                //    rtn = mdcp.加入文書封裝(xnd.Attributes[0].Value.Trim(), xnd.Attributes[1].Value.Trim(), difilever, atts);
                //}
                //else
                //{
                    //string difile = Path.Combine(Path.Combine(CommonTool.CheckoutfilePath, xnd.Attributes[0].Value.Trim()), xnd.Attributes[0].Value.Trim() + ".di");
                    string difilever = xnd.Attributes[0].Value.Trim() + "V" + xnd.Attributes[2].Value.Trim() + ".di";
                    if (!File.Exists(Path.Combine(mdcp.檔案來源資料夾, difilever)))
                    {
                        string minus = xnd.Attributes[0].Value.Trim() + "V" + Convert.ToString(int.Parse(xnd.Attributes[2].Value.Trim()) - 1) + ".di";
                        int cnt = 0;
                        while (!File.Exists(Path.Combine(mdcp.檔案來源資料夾, minus)))
                        {
                            if (int.Parse(xnd.Attributes[2].Value.Trim()) >= (1 + cnt))
                            {
                                minus = xnd.Attributes[0].Value.Trim() + "V" + Convert.ToString(int.Parse(xnd.Attributes[2].Value.Trim()) - (1 + cnt)) + ".di";
                            }
                            else
                                break;
                            cnt++;
                        }
                        File.Copy(Path.Combine(mdcp.檔案來源資料夾, minus), Path.Combine(mdcp.檔案來源資料夾, difilever));
                    }
                    //處理附件
                    List<string> atts = new List<string>();
                    foreach (XmlNode chnd in xnd.ChildNodes)
                    {
                        if (chnd.InnerXml == "") continue;
                        FileInfo sourceAtt = new FileInfo(Path.Combine(sourcefilepath, chnd.Attributes["AtaFile"].Value.Trim()));
                        string targetAtt = Path.Combine(sourcefilepath, sourceAtt.Name.Replace(sourceAtt.Extension, "") + "V" + chnd.Attributes["LastChangeVer"].Value.Trim() + sourceAtt.Extension);
                        atts.Add(sourceAtt.Name.Replace(sourceAtt.Extension, "") + "V" + chnd.Attributes["LastChangeVer"].Value.Trim() + sourceAtt.Extension);
                    }
                    rtn = mdcp.加入文書封裝(xnd.Attributes[0].Value.Trim(), xnd.Attributes[1].Value.Trim(), difilever, atts, isjudge);
                //}
            }
            return rtn;
        }


        private void DeleteUnusedFiles()
        {
            DirectoryInfo di = new DirectoryInfo(sourcefilepath);
            XmlNodeList usefiles = mdcp.封裝檔XML.GetElementsByTagName("檔案名稱");
            List<string> fname = new List<string>();
            foreach (XmlNode nd in usefiles)
                fname.Add(nd.InnerText);
            foreach (FileInfo f in di.GetFiles())
            {
                if (f.Name == DocNO + ".si" || multiOEDocNO.Contains(f.Name.Replace(".tif", ""))) continue;
                if (!fname.Contains(f.Name)) f.Delete();
            }

        }
        #endregion
    }
}
