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
    /// �����x�s���������
    /// </summary>
    /// <param name="xdoc">��������ɤ��e</param>
    /// <param name="sMetaDataPath">��������ɸ��|</param>
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
    /// ���X�C��ʸ���(�������ҥd�A��)
    /// </summary>
    /// <param name="sMetaDataPath">��������ɸ��|</param>
    /// <param name="CertPath">�����������_���|</param>
    /// <param name="sMediaPackagePath">�C��ʸ��ɲ��X���|</param>
    /// <param name="CertPwd">�������ұK�X</param>
    /// <returns></returns>
    [WebMethod]
    public bool WSEMakeEMediaXml(string sMetaDataPath, string CertPath, string sMediaPackagePath, string CertPwd)
    {
        //CertPath ��������������
        //CertPwd �������ұK�X
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
    /// ���X�C��ʸ���(HSM�A��)
    /// </summary>
    /// <param name="sMetaDataPath">��������ɸ��|</param>
    /// <param name="CertPath">�����������_���|</param>
    /// <param name="sMediaPackagePath">�C��ʸ��ɲ��X���|</param>
    /// <param name="sKeyName">HSM�ѧO�X</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEMakeEMediaXmlFormHSM(string sMetaDataPath, string CertPath, string sMediaPackagePath, string sKeyName)
    {

        string srtnHSM = "";    //HSM�[ñ����^�Ǫ���XML�r��
        string rtn = "";    //WebMothed�^�ǭ�

        try
        {
            DCPWebService.DCPWSE dcpwse = new DCPWebService.DCPWSE();
            DigitalSealed.OA.EMedia mediaxml;
            rtn = "�I�s���󲣥X�C��ʸ��ɥ���.";
            if (CertPath.Trim() == "")
                mediaxml = new DigitalSealed.OA.EMedia(sMetaDataPath);
            else
                mediaxml = new DigitalSealed.OA.EMedia(sMetaDataPath, CertPath);
            rtn = "�I�s����N�C��ʸ��ɥt�s����.";
            mediaxml.MakeEMediaXml("").Save(sMediaPackagePath);
            XmlDocument xdoc = new XmlDocument();
            rtn = "LOAD�C��ʸ��ɥ���.";
            xdoc.Load(sMediaPackagePath); //Ū���w�p��hash��xml��
            rtn = "�I�sHSM�[ñ����.";
            srtnHSM = dcpwse.SignWithHSM(xdoc.OuterXml, "�C��ʸ���", sKeyName);
            XmlDocument xdoc_1 = new XmlDocument();
            rtn = "�B�zHSM�^�Ǧr�ꥢ��(LOAD).";
            xdoc_1.XmlResolver = null;
            xdoc_1.LoadXml(srtnHSM);
            rtn = "�B�zHSM�^�Ǧr�ꥢ��(SAVE).";
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
    ///  ���X���_�H��(�A�ξ������ҥd)
    /// </summary>
    /// <param name="sMediaPackagePath">�C��ʸ��ɸ��|</param>
    /// <param name="medianum">�q�l�C��s��</param>
    /// <param name="CertPwd">�������ұK�X</param>
    /// <param name="sTransPackagePath">���_�H�ʲ��X���|</param>
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
    /// ���X���_�H��(�A��HSM)
    /// </summary>
    /// <param name="sMediaPackagePath">�C��ʸ��ɸ��|</param>
    /// <param name="medianum">�q�l�C��s��</param>
    /// <param name="keyname">HSM�ѧO�X</param>
    /// <param name="sTransPackagePath">���_�H�ʲ��X���|</param>
    /// <returns></returns>
    [WebMethod]
    public bool WSEMakeTransferXmlForHSM(string sMediaPackagePath, string[] medianum, string keyname, string sTransPackagePath)
    {
        try
        {
            
            
            DigitalSealed.OA.ETransfer et = new DigitalSealed.OA.ETransfer(sMediaPackagePath);
            string rtndoc = et.MakeTransferXml(medianum, "").OuterXml;
            DCPWebService.DCPWSE dcpwse = new DCPWebService.DCPWSE();
            string rtn = dcpwse.SignWithHSM(rtndoc, et.�����ʸ���SigID, keyname);
            if (rtn.IndexOf("err:") > -1) //�o�Ϳ��~
            {
                return false;
            }
            else
            {
                //File.WriteAllText(sTransPackagePath, rtn,  Encoding.UTF8); //�g�X�̫�[ñ��XML��
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
    /// �����I���m�J��������(�A�ξ������ҥd)
    /// </summary>
    /// <param name="SIFile">�u�Wñ���ɸ��|���ɦW</param>
    /// <param name="XMLFile">�ʸ��ɸ��|���ɦW</param>
    /// <param name="CertPwd">�������ұK�X</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEArchivistOAFileOrgCert(string SIFile, string XMLFile, string CertPwd)
    {
        string rtn = "";
        try
        {
            DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(SIFile, XMLFile);
            rtn = FileOA.�ɺ��I���[��������(CertPwd);
        }
        catch
        {
            return "����";
        }
        return rtn;
    }

    /// <summary>
    /// �����I���m�J��������(�A��HSM)
    /// </summary>
    /// <param name="SIFile">�u�Wñ���ɸ��|���ɦW</param>
    /// <param name="XMLFile">�ʸ��ɸ��|���ɦW</param>
    /// <param name="keyname">HSM�̪���������id</param>
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
            rtn = "�إ�DCP���󥢱�!";
            DigitalSealed.OA.Archivist FileOA = new DigitalSealed.OA.Archivist(SIFile, XMLFile);
            rtn = "�I�sHSM�[ñ����!";
            srtnHSM = dcpwse.SignWithHSM(FileOA.���o�ɺ��I���[ñ�e���(), "CheckSignGCA", keyname);
            if (srtnHSM.StartsWith("err"))
            {
                File.WriteAllText(SIFile + "err.txt", srtnHSM);
                return rtn;
            }

            try
            {
                rtn = "�B�zHSM�^�Ǥ�XML�r�ꥢ��(LoadXml)!";
                xdoc.XmlResolver = null;
                xdoc.LoadXml(srtnHSM);
                rtn = "�B�zHSM�^�Ǥ�XML�r�ꥢ��(Save)!";
                xdoc.Save(XMLFile);
                rtn = "";
            }
            catch (Exception errXML)
            {
                if (rtn.Trim() == "")
                    rtn = "�NHSM�^�Ǥ�XML�r��s��XML�ɵo�ͨҥ~���~!" + errXML.Message;
                else
                    rtn += errXML.Message;
            }
            finally
            {

            }

        }
        catch
        {
            return "����";
        }
        return rtn;

    }

    /// <summary>
    /// �m�J�������(�A�ξ������ҥd)
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
            rtn = FileOA.�[�J�������(NodeRoot, CertPwd);
        }
        catch
        {
            return "����";
        }
        return rtn;
    }

    /// <summary>
    /// �m�J�������(�A��HSM)
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
            rtn = FileOA.�[�J�������(NodeRoot, "");
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
                    rtn = "LOAD�ʸ��ɥ���." + errLoad.Message;
                }
                if (rtn == "")
                {
                    try
                    {
                        rtn = "�I�sHSM�[ñ����.";
                        rtnHSM = dcpwse.SignWithHSM(xdoc.OuterXml, FileOA.�������SignID, KeyName);
                        XmlDocument xdoc_1 = new XmlDocument();
                        xdoc_1.XmlResolver = null;
                        rtn = "�B�zHSM�^�Ǧr�ꥢ��(LoadXml).";
                        xdoc_1.LoadXml(rtnHSM);
                        xdoc_1.PreserveWhitespace = true;
                        rtn = "�B�zHSM�^�Ǧr�ꥢ��(Save).";
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
                rtn = "�[�J������ƥ���!DCP����^�ǰT��-" + rtn + "";
            }
        }
        catch (Exception err)
        {
            return "����" + err.Message;
        }
        return rtn;
    }

    /// <summary>
    /// ���X�v���ʸ���(���t�������)
    /// </summary>
    /// <param name="sDocNO">����帹</param>
    /// <param name="sDCFilePath">�v���ʸ��ɸ��|</param>
    /// <param name="sWorkAreaPath">�v���ӷ����|</param>
    /// <returns></returns>
    [WebMethod]
    public string WSEMakeMediaPackage(string sDocNO, string sDCFilePath, string sWorkAreaPath)
    {
        //�v���ʸ�
        string rtn = "";
        try
        {
            DigitalSealed.OA.MediaPackage mp = new DigitalSealed.OA.MediaPackage(sDocNO, sWorkAreaPath);
            rtn = mp.MakeMediaPackage(sDCFilePath);
        }
        catch
        {
            return "����";
        }
        return rtn;
    }


    /// <summary>
    /// �ˬd�ʸ��ɬO�_�L�~
    /// </summary>
    /// <param name="sDocElecPath">�q�l�ʸ��ɸ��|</param>
    /// <param name="siFileNam">�q�l�ʸ����ɦW</param>
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
            return "�M�d����-" + err.Message;
        }
        
    }

    /// <summary>
    /// ���o�u�Wñ�֤��孶���`�M
    /// </summary>
    /// <param name="sDocElecFileFullPath">�q�l�ʸ��ɧ�����|</param>
    /// <returns></returns>
    [WebMethod]
    public string GetSiDocPage(string sDocElecFileFullPath)
    {
        int iPage = 0;
        try
        {
            if (File.Exists(sDocElecFileFullPath))
            {
                #region ���u�Wñ�֭���
                XmlDocument xdoc = new XmlDocument();
                xdoc.XmlResolver = null;
                xdoc.Load(sDocElecFileFullPath);
                XmlNodeList xnodeRoot = xdoc.SelectNodes("�u�Wñ��");
                if (xnodeRoot.Count > 0)
                {
                    foreach (XmlNode xNode_1 in xnodeRoot)
                    {
                        foreach (XmlNode xNode_2 in xNode_1.ChildNodes)
                        {
                            if (xNode_2.Name == "�u�Wñ�ָ�T")
                            {
                                #region �u�Wñ�ָ�T
                                foreach (XmlNode xNode_3 in xNode_2.ChildNodes)
                                {
                                    #region ��ñ���I�F
                                    if (xNode_3.Attributes.Item(0).Value.IndexOf("_�M��") > -1)
                                    {
                                        foreach (XmlNode xNode_4 in xNode_3.ChildNodes)
                                        {
                                            if (xNode_4.Name == "Object")
                                            {
                                                foreach (XmlNode xNode_5 in xNode_4.ChildNodes)
                                                {
                                                    if (xNode_5.Name == "ñ�ָ�T")
                                                    {
                                                        foreach (XmlNode xNode_6 in xNode_5.ChildNodes)
                                                        {
                                                            if (xNode_6.Name == "ñ�֤��")
                                                            {
                                                                foreach (XmlNode xNode_7 in xNode_6.ChildNodes)
                                                                {
                                                                    if (xNode_7.Name == "ñ�֤�Z�M��")
                                                                    {
                                                                        foreach (XmlNode xNode_8 in xNode_7.ChildNodes)
                                                                        {
                                                                            if (xNode_8.Name == "��Z")
                                                                            {
                                                                                foreach (XmlNode xNode_9 in xNode_8.ChildNodes)
                                                                                {
                                                                                    if (xNode_9.Name == "��Z������")
                                                                                    {
                                                                                        foreach (XmlNode xNode_10 in xNode_9.ChildNodes)
                                                                                        {
                                                                                            if (xNode_10.Name == "��Z�����M��")
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