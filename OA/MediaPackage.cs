using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using DigitalSealed.Tools;
using System.Drawing;
using System.Drawing.Imaging;

namespace DigitalSealed.OA
{
    /// <summary>
    /// 影音封裝類別
    /// </summary>
    public class MediaPackage
    {
        private string _docno = "";
        private DirectoryInfo _DocIMG = null;
        private DirectoryInfo _AttIMG = null;
        private List<string> refs = new List<string>();
        private XmlNode _其他電子影音檔 = null;
        private string _outputpath = "";

        public MediaPackage() { }
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="docno">文號</param>
        /// <param name="docimg">本文影像檔資料夾位置</param>
        /// <param name="attimg">附件影像檔資料夾位置</param>
        public MediaPackage(string docno,string docimg,string attimg)
        {
            _docno = docno;
            _DocIMG = new DirectoryInfo(docimg);
            _AttIMG = new DirectoryInfo(attimg); 
        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="docno">文號</param>
        /// <param name="docimg">本文影像檔資料夾位置</param>
        public MediaPackage(string docno, string docimg)
        {
            _docno = docno;
            _DocIMG = new DirectoryInfo(docimg);
        }

        /// <summary>
        /// 產生影音封裝檔
        /// </summary>
        /// <param name="descdata">詮釋資料</param>
        /// <param name="pwd">憑證密碼</param>
        /// <param name="outputFilePath">輸出檔的路徑(預設檔名:文號.xml)</param>
        /// <returns>成功回傳空字串</returns>
        public string MakeMediaPackage(XmlNode descdata, string pwd, string outputFilePath)
        {
            string rtn = "";
            try
            {
                DirectoryInfo difo = new DirectoryInfo(outputFilePath);
                if (!difo.Exists) difo.Create();
                _outputpath = outputFilePath;
                newMediaDcpXmlFile().Save(Path.Combine(outputFilePath, _docno + ".xml"));
                Archivist arch = new Archivist(Path.Combine(outputFilePath, _docno + ".xml"), refs);
                rtn = arch.加入詮釋資料(descdata, pwd);
            }
            catch (Exception err)
            {
                rtn = err.Message;
            }
            return rtn;
        }

        public string MakeMediaPackage(string outputFilePath)
        {
            string rtn = "";
            try
            {
                DirectoryInfo difo = new DirectoryInfo(outputFilePath);
                if (!difo.Exists) difo.Create();
                _outputpath = outputFilePath;
                XmlDocument doc = newMediaDcpXmlFile();
                string fullpath = Path.Combine(outputFilePath, _docno + ".xml");
                doc.Save(fullpath);
                Archivist arch = new Archivist(fullpath, refs);
                rtn = arch.加入詮釋資料(null, "");
            }
            catch (Exception err)
            {
                rtn = err.Message;
            }
            return rtn;
        }

        /// <summary>
        /// 產生影音封裝檔
        /// </summary>
        /// <param name="descdata">詮釋資料</param>
        /// <param name="otherdata">其它電子影音檔</param>
        /// <param name="pwd">憑證密碼</param>
        /// <param name="outputFilePath">輸出檔的路徑(預設檔名:文號.xml)</param>
        /// <returns>成功回傳空字串</returns>
        public string MakeMediaPackage(XmlNode descdata,XmlNode otherdata, string pwd, string outputFilePath)
        {
            string rtn = "";
            try
            {
                DirectoryInfo difo = new DirectoryInfo(outputFilePath);
                if (!difo.Exists) difo.Create();
                _outputpath = outputFilePath;
                _其他電子影音檔 = otherdata;
                newMediaDcpXmlFile().Save(Path.Combine(outputFilePath, _docno + ".xml"));
                Archivist arch = new Archivist(Path.Combine(outputFilePath, _docno + ".xml"), refs);
                arch.加入詮釋資料(descdata, pwd);
            }
            catch (Exception err)
            {
                rtn = err.Message;
            }
            return rtn;
        }

        /// <summary>
        /// 產生新的電子封裝檔
        /// </summary>
        private XmlDocument newMediaDcpXmlFile()
        {
            XmlDocument DcpXml = new XmlDocument();
            DcpXml.XmlResolver = null;
            DcpXml.PreserveWhitespace = true;
            DcpXml.AppendChild(DcpXml.CreateXmlDeclaration("1.0", "utf-8", null));
            DcpXml.AppendChild(DcpXml.CreateDocumentType("電子封裝檔", null, "99_erencaps_utf8.dtd", null));

            XmlNode 封裝檔內容 = xmlTool.MakeNode("封裝檔內容", "Id", "Wrap");
            封裝檔內容.AppendChild(xmlTool.MakeNode("封裝檔資訊", "電子影音檔案"));
            XmlNode 電子影音檔案 = xmlTool.MakeNode("電子影音檔案", "文號", _docno);
            addFileInfo(ref 電子影音檔案);
            封裝檔內容.AppendChild(電子影音檔案);
            XmlNode 電子封裝檔 = xmlTool.MakeNode("電子封裝檔", "");
            電子封裝檔.AppendChild(封裝檔內容);
            DcpXml.AppendChild(DcpXml.ImportNode(電子封裝檔, true));

            return DcpXml;
        }

        /// <summary>
        /// 產生歸檔掃描影像
        /// </summary>
        /// <param name="電子影音檔案"></param>
        private void addFileInfo(ref XmlNode 電子影音檔案)
        {
            Dictionary<string, string> atts = new Dictionary<string, string>();

            int _filecnt = 0;
            foreach (FileInfo file in _DocIMG.GetFiles())
            {
                if (file.Attributes == FileAttributes.Normal || file.Attributes == FileAttributes.Archive)
                { _filecnt++; }
            }
            int _pagecnt = 0;
            atts.Add("頁數", "");
            atts.Add("檔案數", _filecnt.ToString());
            atts.Add("群組名稱", "本文");
            atts.Add("群組型別", "本文");
            XmlNode 本文頁面群組 = xmlTool.MakeNode("頁面群組", atts);
             
            foreach (FileInfo file in _DocIMG.GetFiles())
            {
                if (file.Attributes == FileAttributes.Normal || file.Attributes == FileAttributes.Archive)
                {
                    XmlNode mediafile = xmlTool.MakeNode("電子影音檔案資訊", "");
                    mediafile.AppendChild(xmlTool.MakeNode("檔案名稱", file.Name));
                    refs.Add(file.Name);
                    mediafile.AppendChild(xmlTool.MakeNode("檔案大小", file.Length.ToString()));
                    string fmt = (file.Extension.StartsWith(".") ? file.Extension.Substring(1) : file.Extension).ToUpper();
                    mediafile.AppendChild(xmlTool.MakeNode("檔案格式", fmt));
                    本文頁面群組.AppendChild(mediafile);
                    if (File.Exists(Path.Combine(_outputpath, file.Name))) File.Delete(Path.Combine(_outputpath, file.Name));
                    file.CopyTo(Path.Combine(_outputpath, file.Name), true);
                    if (file.Extension.ToLower().IndexOf("tif") > -1)
                    {
                        _pagecnt += PageCount(Image.FromFile(file.FullName));
                    }
                    if (file.Extension.ToLower().IndexOf("pdf") > -1)
                    {
                        _pagecnt += pdfPageCount(file.FullName);
                    }
                }
            }
            本文頁面群組.Attributes["頁數"].Value = _pagecnt.ToString();
            
            int _attfilecnt = 0;
            int _attpagecnt = 0;
            XmlNode 附件頁面群組 = null;
            if (_AttIMG != null)
            {
                atts.Clear();
                _attfilecnt = _AttIMG.GetFiles().Length;
                atts.Add("頁數", "");
                atts.Add("檔案數", _attfilecnt.ToString());
                atts.Add("群組名稱", "附件");
                atts.Add("群組型別", "附件");
                附件頁面群組 = xmlTool.MakeNode("頁面群組", atts);
                foreach (FileInfo file in _AttIMG.GetFiles())
                {
                    XmlNode mediafile = xmlTool.MakeNode("電子影音檔案資訊", "");
                    mediafile.AppendChild(xmlTool.MakeNode("檔案名稱", file.Name));
                    refs.Add(file.Name);
                    mediafile.AppendChild(xmlTool.MakeNode("檔案大小", file.Length.ToString()));
                    string fmt = (file.Extension.StartsWith(".") ? file.Extension.Substring(1) : file.Extension).ToUpper();
                    mediafile.AppendChild(xmlTool.MakeNode("檔案格式", fmt));
                    附件頁面群組.AppendChild(mediafile);
                    file.CopyTo(Path.Combine(_outputpath, file.Name),true);
                    if (file.Extension.ToLower().IndexOf("tif") > -1)
                    {
                        _attpagecnt += PageCount(Image.FromFile(file.FullName));
                    }
                    if (file.Extension.ToLower().IndexOf("pdf") > -1)
                    {
                        _attpagecnt += pdfPageCount(file.FullName);
                    }
                }
                附件頁面群組.Attributes["頁數"].Value = _attpagecnt.ToString();
            }

            atts.Clear();
            atts.Add("群組數", _AttIMG == null ? "1" : "2");
            atts.Add("總頁數", (_pagecnt + _attpagecnt).ToString());
            XmlNode 歸檔掃描影像 = xmlTool.MakeNode("歸檔掃描影像", atts);
            歸檔掃描影像.AppendChild(本文頁面群組);
            if (附件頁面群組 != null)
                歸檔掃描影像.AppendChild(附件頁面群組);
            電子影音檔案.AppendChild(歸檔掃描影像);
            if (_其他電子影音檔 != null)
            {
                foreach (XmlNode nd in _其他電子影音檔.SelectNodes("//影音檔群組/電子影音檔案資訊/檔案名稱"))
                {
                    refs.Add(nd.InnerText); 
                }
                電子影音檔案.AppendChild(_其他電子影音檔);
            }
        }

        private int pdfPageCount(string pdffile)
        {
            string pdfText = "";
            using (FileStream fs = new FileStream(pdffile, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader r = new StreamReader(fs))
                {
                    pdfText = r.ReadToEnd();
                }
            }
            System.Text.RegularExpressions.Regex rx1 = new System.Text.RegularExpressions.Regex(@"/Type\s*/Page[^s]");
            System.Text.RegularExpressions.MatchCollection matches = rx1.Matches(pdfText);
            return matches.Count;
        }

        private int PageCount(Image img)
        {
            int pageCount = -1;
            try
            {
                pageCount = img.GetFrameCount(FrameDimension.Page);
                img.Dispose(); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return pageCount;
        }

        private Image[] SplitImages(Image MergedImage, int pageCount)
        {
            MemoryStream ms = null;

            Image[] SplitImages = new Image[pageCount];

            try
            {
                Guid objGuid = MergedImage.FrameDimensionsList[0];
                FrameDimension objDimension = new FrameDimension(objGuid);

                for (int i = 0; i < pageCount; i++)
                {
                    ms = new MemoryStream();
                    MergedImage.SelectActiveFrame(objDimension, i);
                    MergedImage.Save(ms, ImageFormat.Tiff);
                    SplitImages[i] = Image.FromStream(ms);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.alert(ex.ToString());
            }
            finally
            {
                ms.Close();
            }

            return SplitImages;
        }
    }
}
