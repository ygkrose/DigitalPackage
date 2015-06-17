<%@ WebService Language="C#" Class="WSEEMediaXml" %>

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.IO;
using System.Xml;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class WSEEMediaXml : System.Web.Services.WebService
{
    /// <summary>
    /// 遠端儲存詮釋資料檔
    /// </summary>
    /// <param name="xdoc">詮釋資料檔內容</param>
    /// <param name="sMetaDataPath">詮釋資料檔路徑</param>
    /// <returns></returns>
    [WebMethod]
    public bool WSEXMLSave(XmlDocument xdoc, string sMetaDataPath)
    {
        System.IO.StreamWriter txt = null;
        try
        {

            txt = new System.IO.StreamWriter(sMetaDataPath, false, System.Text.Encoding.GetEncoding("UTF-8"));
            XmlDocument xdoc_1 = new XmlDocument();
            try
            {
                XmlNode xnode = xdoc.SelectSingleNode("ROWSET");
                xnode.Attributes.RemoveAll();
                txt.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                txt.WriteLine("<!DOCTYPE ROWSET SYSTEM \"99_ermeta_utf8.dtd\">");
                txt.WriteLine(xdoc.InnerXml);


            }
            catch
            {

            }
            //       xdoc.Save(sMetaDataPath);

        }
        catch
        {
            return false;
        }
        finally
        {
            txt.Flush();
            txt.Close();
        }
        return true;
    }


    /// <summary>
    /// 產出媒體封裝檔(機關憑證卡適用)
    /// </summary>
    /// <param name="sMetaDataPath">詮釋資料檔路徑</param>
    /// <param name="CertPath">接關機關公鑰路徑</param>
    /// <param name="sMediaPackagePath">媒體封裝檔產出路徑</param>
    /// <param name="CertPwd">機關憑證密碼</param>
    /// <returns></returns>
    [WebMethod]
    public bool WSEMakeEMediaXml(string sMetaDataPath, string CertPath, string sMediaPackagePath, string CertPwd)
    {
        //CertPath 接受機關憑證檔
        //CertPwd 機關憑證密碼
        try
        {
            DigitalSealed.OA.EMedia mediaxml;
            if (CertPath.Trim() == "")
                mediaxml = new DigitalSealed.OA.EMedia(sMetaDataPath);
            else
                mediaxml = new DigitalSealed.OA.EMedia(sMetaDataPath, CertPath);

            mediaxml.MakeEMediaXml(CertPwd).Save(sMediaPackagePath);
        }
        catch
        {
            return false;
        }
        return true;
    }

    /// <summary>
    /// 產出媒體封裝檔(HSM適用)
    /// </summary>
    /// <param name="sMetaDataPath">詮釋資料檔路徑</param>
    /// <param name="CertPath">接關機關公鑰路徑</param>
    /// <param name="sMediaPackagePath">媒體封裝檔產出路徑</param>
    /// <param name="sKeyName">HSM識別碼</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEMakeEMediaXmlFormHSM(string sMetaDataPath, string CertPath, string sMediaPackagePath, string sKeyName)
    {

        string srtnHSM = "";    //HSM加簽完後回傳的為XML字串
        string rtn = "";    //WebMothed回傳值

        try
        {
            DCPWebService.DCPWSE dcpwse = new DCPWebService.DCPWSE();
            DigitalSealed.OA.EMedia mediaxml;
            rtn = "呼叫元件產出媒體封裝檔失敗.";
            if (CertPath.Trim() == "")
                mediaxml = new DigitalSealed.OA.EMedia(sMetaDataPath);
            else
                mediaxml = new DigitalSealed.OA.EMedia(sMetaDataPath, CertPath);
            rtn = "呼叫元件將媒體封裝檔另存失敗.";
            mediaxml.MakeEMediaXml("").Save(sMediaPackagePath);
            XmlDocument xdoc = new XmlDocument();
            rtn = "LOAD媒體封裝檔失敗.";
            xdoc.Load(sMediaPackagePath); //讀取已計算hash的xml檔
            rtn = "呼叫HSM加簽失敗.";
            srtnHSM = dcpwse.SignWithHSM(xdoc.OuterXml, "媒體封裝檔", sKeyName);
            XmlDocument xdoc_1 = new XmlDocument();
            rtn = "處理HSM回傳字串失敗(LOAD).";
            xdoc_1.XmlResolver = null;
            xdoc_1.LoadXml(srtnHSM);
            rtn = "處理HSM回傳字串失敗(SAVE).";
            xdoc_1.Save(sMediaPackagePath);
            rtn = "";
            return rtn;


        }
        catch (Exception err)
        {
            return rtn + " " + err.Message;
        }

    }

    /// <summary>
    ///  產出金鑰信封(適用機關憑證卡)
    /// </summary>
    /// <param name="sMediaPackagePath">媒體封裝檔路徑</param>
    /// <param name="medianum">電子媒體編號</param>
    /// <param name="CertPwd">機關憑證密碼</param>
    /// <param name="sTransPackagePath">金鑰信封產出路徑</param>
    /// <returns></returns>
    [WebMethod]
    public bool WSEMakeTransferXml(string sMediaPackagePath, string[] medianum, string CertPwd, string sTransPackagePath)
    {
        try
        {
            DigitalSealed.OA.ETransfer et = new DigitalSealed.OA.ETransfer(sMediaPackagePath);
            XmlDocument rtndoc = et.MakeTransferXml(medianum, CertPwd);
            rtndoc.Save(sTransPackagePath);
        }
        catch
        {
            return false;
        }
        return true;
    }
    /// <summary>
    /// 產出金鑰信封(適用HSM)
    /// </summary>
    /// <param name="sMediaPackagePath">媒體封裝檔路徑</param>
    /// <param name="medianum">電子媒體編號</param>
    /// <param name="keyname">HSM識別碼</param>
    /// <param name="sTransPackagePath">金鑰信封產出路徑</param>
    /// <returns></returns>
    [WebMethod]
    public bool WSEMakeTransferXmlForHSM(string sMediaPackagePath, string[] medianum, string keyname, string sTransPackagePath)
    {
        try
        {
            
            
            DigitalSealed.OA.ETransfer et = new DigitalSealed.OA.ETransfer(sMediaPackagePath);
            string rtndoc = et.MakeTransferXml(medianum, "").OuterXml;
            DCPWebService.DCPWSE dcpwse = new DCPWebService.DCPWSE();
            string rtn = dcpwse.SignWithHSM(rtndoc, et.移轉交封裝檔SigID, keyname);
            if (rtn.IndexOf("err:") > -1) //發生錯誤
            {
                return false;
            }
            else
            {
                //File.WriteAllText(sTransPackagePath, rtn,  Encoding.UTF8); //寫出最後加簽的XML檔
                XmlDocument xdoc_1 = new XmlDocument();
                xdoc_1.XmlResolver = null;
                xdoc_1.LoadXml(rtn);
                xdoc_1.PreserveWhitespace = true;
                xdoc_1.Save(sTransPackagePath);                
            }

        }
        catch
        {
            return false;
        }
        return true;
    }
    /// <summary>
    /// 公文點收置入機關憑證(適用機關憑證卡)
    /// </summary>
    /// <param name="SIFile">線上簽核檔路徑及檔名</param>
    /// <param name="XMLFile">封裝檔路徑及檔名</param>
    /// <param name="CertPwd">機關憑證密碼</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEArchivistOAFileOrgCert(string SIFile, string XMLFile, string CertPwd)
    {
        string rtn = "";
        try
        {
            DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(SIFile, XMLFile);
            rtn = FileOA.檔管點收加機關憑證(CertPwd);
        }
        catch
        {
            return "失敗";
        }
        return rtn;
    }

    /// <summary>
    /// 公文點收置入機關憑證(適用HSM)
    /// </summary>
    /// <param name="SIFile">線上簽核檔路徑及檔名</param>
    /// <param name="XMLFile">封裝檔路徑及檔名</param>
    /// <param name="keyname">HSM裡的機關憑證id</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEArchivistOAFileOrgCertForHSM(string SIFile, string XMLFile, string keyname)
    {
        string rtn = "";
        string srtnHSM = "";
        XmlDataDocument xdoc = new XmlDataDocument();

        try
        {
            DCPWebService.DCPWSE dcpwse = new DCPWebService.DCPWSE();
            rtn = "建立DCP元件失敗!";
            DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(SIFile, XMLFile);
            rtn = "呼叫HSM加簽失敗!";
            srtnHSM = dcpwse.SignWithHSM(FileOA.取得檔管點收加簽前資料(), "CheckSignGCA", keyname);
            if (srtnHSM.StartsWith("err"))
            {
                File.WriteAllText(SIFile + "err.txt", srtnHSM);
                return rtn;
            }

            try
            {
                rtn = "處理HSM回傳之XML字串失敗(LoadXml)!";
                xdoc.XmlResolver = null;
                xdoc.LoadXml(srtnHSM);
                rtn = "處理HSM回傳之XML字串失敗(Save)!";
                xdoc.Save(XMLFile);
                rtn = "";
            }
            catch (Exception errXML)
            {
                if (rtn.Trim() == "")
                    rtn = "將HSM回傳之XML字串存成XML時發生例外錯誤!" + errXML.Message;
                else
                    rtn += errXML.Message;
            }
            finally
            {

            }

        }
        catch
        {
            return "失敗";
        }
        return rtn;

    }

    /// <summary>
    /// 置入詮釋資料(適用機關憑證卡)
    /// </summary>
    /// <param name="NodeRoot"></param>
    /// <param name="XMLFile"></param>
    /// <param name="CertPwd"></param>
    /// <returns></returns>
    [WebMethod]
    public string WSEExplainData(XmlNode NodeRoot, string XMLFile, string CertPwd)
    {
        string rtn = "";
        try
        {
            NodeRoot.Attributes.RemoveAll();
            DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(XMLFile);
            rtn = FileOA.加入詮釋資料(NodeRoot, CertPwd);
        }
        catch
        {
            return "失敗";
        }
        return rtn;
    }

    /// <summary>
    /// 置入詮釋資料(適用HSM)
    /// </summary>
    /// <param name="NodeRoot"></param>
    /// <param name="XMLFile"></param>
    /// <param name="KeyName"></param>
    /// <returns></returns>

    [WebMethod]
    public string WSEExplainDataForHSM(XmlNode NodeRoot, string XMLFile, string KeyName, string sElecPath, string sDocNO)
    {
        string rtn = "";
        string rtnHSM = "";
        XmlDocument xdoc = new XmlDocument();
        try
        {
            DCPWebService.DCPWSE dcpwse = new DCPWebService.DCPWSE();

            NodeRoot.Attributes.RemoveAll();
            //DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(XMLFile);
            DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(XMLFile, sElecPath + sDocNO + "_MetaData.xml");
            rtn = FileOA.加入詮釋資料(NodeRoot, "");
            if (rtn == "")
            {
                try
                {
                    xdoc.XmlResolver = null;
                    //xdoc.Load(XMLFile);                    
                    xdoc.Load(sElecPath + sDocNO + "_MetaData.xml");
                    xdoc.PreserveWhitespace = true;
                }
                catch (Exception errLoad)
                {
                    rtn = "LOAD封裝檔失敗." + errLoad.Message;
                }
                if (rtn == "")
                {
                    try
                    {
                        rtn = "呼叫HSM加簽失敗.";
                        rtnHSM = dcpwse.SignWithHSM(xdoc.OuterXml, FileOA.詮釋資料SignID, KeyName);
                        XmlDocument xdoc_1 = new XmlDocument();
                        xdoc_1.XmlResolver = null;
                        rtn = "處理HSM回傳字串失敗(LoadXml).";
                        xdoc_1.LoadXml(rtnHSM);
                        xdoc_1.PreserveWhitespace = true;
                        rtn = "處理HSM回傳字串失敗(Save).";
                        xdoc_1.Save(XMLFile);
                        rtn = "";
                    }
                    catch (Exception errLoad)
                    {
                        if (rtn.Trim() != "")
                            rtn += errLoad.Message;
                        else
                            rtn = errLoad.Message;
                    }

                }
            }
            else
            {
                rtn = "加入詮釋資料失敗!DCP元件回傳訊息-" + rtn + "";
            }
        }
        catch (Exception err)
        {
            return "失敗" + err.Message;
        }
        return rtn;
    }

    /// <summary>
    /// 產出影音封裝檔(不含詮釋資料)
    /// </summary>
    /// <param name="sDocNO">公文文號</param>
    /// <param name="sDCFilePath">影音封裝檔路徑</param>
    /// <param name="sWorkAreaPath">影像來源路徑</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEMakeMediaPackage(string sDocNO, string sDCFilePath, string sWorkAreaPath)
    {
        //影音封裝
        string rtn = "";
        try
        {
            DigitalSealed.OA.MediaPackage mp = new DigitalSealed.OA.MediaPackage(sDocNO, sWorkAreaPath);
            rtn = mp.MakeMediaPackage(sDCFilePath);
        }
        catch
        {
            return "失敗";
        }
        return rtn;
    }


    /// <summary>
    /// 檢查封裝檔是否無誤
    /// </summary>
    /// <param name="sDocElecPath">電子封裝檔路徑</param>
    /// <param name="siFileNam">電子封裝檔檔名</param>
    /// <returns></returns>
    [WebMethod]
    public string WSECheckDigitalSealedSignature(string sDocElecPath, string siFileNam)
    {
        string sErrMsg = "";
        try
        {
            DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature verifysig = new DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature.Signature(sDocElecPath, siFileNam);

            System.Collections.Generic.List<string> rtn = verifysig.VerifyXmlFile();
            if (rtn.Count == 0)
                return "";
            else
            {
                for (int i = 0; i < rtn.Count; i++)
                {
                    sErrMsg += rtn[i].ToString().Trim() + ",";
                }
                return sErrMsg;
            }
        }
        catch (Exception err)
        {
            return "清查失敗-" + err.Message;
        }
        
    }

    /// <summary>
    /// 取得線上簽核公文頁數總和
    /// </summary>
    /// <param name="sDocElecFileFullPath">電子封裝檔完整路徑</param>
    /// <returns></returns>
    [WebMethod]
    public string GetSiDocPage(string sDocElecFileFullPath)
    {
        int iPage = 0;
        try
        {
            if (File.Exists(sDocElecFileFullPath))
            {
                #region 取線上簽核頁數
                XmlDocument xdoc = new XmlDocument();
                xdoc.XmlResolver = null;
                xdoc.Load(sDocElecFileFullPath);
                XmlNodeList xnodeRoot = xdoc.SelectNodes("線上簽核");
                if (xnodeRoot.Count > 0)
                {
                    foreach (XmlNode xNode_1 in xnodeRoot)
                    {
                        foreach (XmlNode xNode_2 in xNode_1.ChildNodes)
                        {
                            if (xNode_2.Name == "線上簽核資訊")
                            {
                                #region 線上簽核資訊
                                foreach (XmlNode xNode_3 in xNode_2.ChildNodes)
                                {
                                    #region 取簽核點了
                                    if (xNode_3.Attributes.Item(0).Value.IndexOf("_決行") > -1)
                                    {
                                        foreach (XmlNode xNode_4 in xNode_3.ChildNodes)
                                        {
                                            if (xNode_4.Name == "Object")
                                            {
                                                foreach (XmlNode xNode_5 in xNode_4.ChildNodes)
                                                {
                                                    if (xNode_5.Name == "簽核資訊")
                                                    {
                                                        foreach (XmlNode xNode_6 in xNode_5.ChildNodes)
                                                        {
                                                            if (xNode_6.Name == "簽核文件夾")
                                                            {
                                                                foreach (XmlNode xNode_7 in xNode_6.ChildNodes)
                                                                {
                                                                    if (xNode_7.Name == "簽核文稿清單")
                                                                    {
                                                                        foreach (XmlNode xNode_8 in xNode_7.ChildNodes)
                                                                        {
                                                                            if (xNode_8.Name == "文稿")
                                                                            {
                                                                                foreach (XmlNode xNode_9 in xNode_8.ChildNodes)
                                                                                {
                                                                                    if (xNode_9.Name == "文稿頁面檔")
                                                                                    {
                                                                                        foreach (XmlNode xNode_10 in xNode_9.ChildNodes)
                                                                                        {
                                                                                            if (xNode_10.Name == "文稿頁面清單")
                                                                                            {
                                                                                                try
                                                                                                {
                                                                                                    iPage += Convert.ToInt16(xNode_10.Attributes.Item(0).Value);
                                                                                                }
                                                                                                catch { }
                                                                                            }

                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                        }



                    }
                }
                #endregion

            }
            return iPage.ToString().Trim();
        }
        catch (Exception errPage)
        {
            return "0" + errPage.Message;

        }

    }
    
    
}