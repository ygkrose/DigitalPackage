using System;
using System.Data;
using System.Configuration;
using System.Diagnostics;
using System.Xml;
using System.IO;
using System.Xml.Xsl;
using System.Drawing;
using System.Drawing.Imaging;

// 引用 WebOE Display 元件
//using Dsic.WebOE.Display;

// 引用 VeryPDF docPrint SDK
using DOCPRINTCOMLib;

namespace DigitalSealed.Tools
{
    /// <summary>
    /// 公文附件文件檔轉影像頁面檔
    /// </summary>
    public class ImageConverter
    {
        /// <summary>
        /// 設定將轉換的影像檔格式。
        /// </summary>
        public enum ImageFileFormat
        { 
            /// <summary>
            /// TIFF
            /// </summary>
            TIF,
            /// <summary>
            /// JPEG
            /// </summary>
            JPG
        }

        /// <summary>
        /// 設定將轉換的影像檔格式。
        /// </summary>
        public enum ConversionComponent
        {
            /// <summary>
            /// docPrint Pro Command Line (doc2pdf.exe)
            /// </summary>
            veryPDF_doc2pdf,
            /// <summary>
            /// docPrint SDK (docPrintCom.dll)
            /// </summary>
            veryPDF_docPrintCom,
            /// <summary>
            /// BCL easyPDF 6.0
            /// </summary>
            easyPDF
        }

        /// <summary>
        /// 錯誤訊息。
        /// </summary>
        public string strLastErrorMessage;

        /// <summary>
        /// 轉換的影像檔資訊。
        /// </summary>
        public FileInfo[] fileInfoOutputImageFiles;

        /// <summary>
        /// 執行文件檔轉影像檔。
        /// </summary>
        /// <param name="strDocumentFilePath">來源文件檔路徑。</param>
        /// <param name="strDocumentFileName">來源文件檔名稱。</param>
        /// <param name="strImageFilePath">轉出頁面影像檔路徑。</param>
        /// <param name="strWordDocNo">文書編號。</param>
        /// <param name="intAttachmentSeqNo">附件檔序號。</param>
        /// <param name="strId_OPDefinition">簽核點定義Id。</param>
        /// <param name="enumImageFileFormat">影像檔格式。</param>
        /// <param name="intImageFileDpi">影像檔 Dpi值。</param>
        /// <param name="longImageFileQuality">影像檔 Quality值。</param>
        /// <param name="enumConversionComponent">轉影像檔元件。</param>
        /// <param name="IsSignaturePage">將轉換的檔是否為簽章資訊頁面檔。</param>
        /// <returns></returns>
        public bool ExecuteDocumentFileConvertToImageFile(
            string strDocumentFilePath, string strDocumentFileName, string strImageFilePath,
            string strWordDocNo, int intAttachmentSeqNo, string strId_OPDefinition,
            ImageFileFormat enumImageFileFormat, int intImageFileDpi, long longImageFileQuality,
            bool IsSignaturePage, ConversionComponent enumConversionComponent)
        {
            try
            {
                // 若 轉出頁面影像檔路徑 不存在則新增
                if (!Directory.Exists(strImageFilePath))
                {
                    Directory.CreateDirectory(strImageFilePath);
                }

                FileInfo fileInfoDocumentFile = new FileInfo(Path.Combine(strDocumentFilePath, strDocumentFileName));

                switch (fileInfoDocumentFile.Extension.Replace(".", "").ToUpper())
                {
                    case "DI":

                        #region 副檔名為 DI

                        if (intAttachmentSeqNo == 0)
                        {
                            // intAttachmentSeqNo == 0 為 電子交換檔
                            if (!ExportImageFile(PreviewDiFileToXmlDocument(fileInfoDocumentFile.FullName),
                                                Path.Combine(strImageFilePath, strWordDocNo + "_W_" + strId_OPDefinition),
                                                enumImageFileFormat, intImageFileDpi, longImageFileQuality))
                            {
                                return false;
                            }
                        }
                        else
                        {
                            // intAttachmentSeqNo > 0 為 附件
                            if (!ExportImageFile(PreviewDiFileToXmlDocument(fileInfoDocumentFile.FullName),
                                                Path.Combine(strImageFilePath,
                                                                strWordDocNo + "_A" + intAttachmentSeqNo.ToString("00")
                                                                             + "_" + strId_OPDefinition),
                                                enumImageFileFormat, intImageFileDpi, longImageFileQuality))
                            {
                                return false;
                            }
                        }

                        #endregion

                        break;                    
                    default:

                        #region 其他非 DI 檔，以 3rd Party Component 轉換

                        if (intAttachmentSeqNo == 0)
                        {
                            string strPageType = (IsSignaturePage) ? "WS" : "W";

                            // intAttachmentSeqNo == 0 為 文稿
                            switch (enumConversionComponent)
                            {
                                case ConversionComponent.veryPDF_doc2pdf:
                                
                                    if (!Execute_veryPDF_doc2pdf(strdocPrintPathName,
                                                     @"-i " + "\"" + fileInfoDocumentFile.FullName + "\" "
                                                   + @"-o " + "\"" + Path.Combine(strImageFilePath,
                                                                            strWordDocNo + "_" + strPageType
                                                                                         + "_" + strId_OPDefinition
                                                                                         + "." + enumImageFileFormat.ToString()) + "\" "
                                                   + @"-d -b 24 -r " + intImageFileDpi.ToString() + " -D -E 2"))
                                    {
                                        return false;
                                    }

                                    break;

                                case ConversionComponent.veryPDF_docPrintCom:

                                    if (!Execute_veryPDF_docPrintCom("", "", fileInfoDocumentFile.FullName,
                                                                Path.Combine(strImageFilePath,
                                                                                strWordDocNo + "_" + strPageType
                                                                                             + "_" + strId_OPDefinition
                                                                                             + "." + enumImageFileFormat.ToString()),
                                                                @"-d -b 24 -r " + intImageFileDpi.ToString() + " -D -E 2"))
                                    {
                                        return false;
                                    }
                                    break;
                            
                                case ConversionComponent.easyPDF:

                                    if (!Execute_easyPDF(fileInfoDocumentFile.FullName,
                                                            Path.Combine(strImageFilePath,
                                                                            strWordDocNo + "_" + strPageType
                                                                                         + "_" + strId_OPDefinition
                                                                                         + "." + enumImageFileFormat.ToString()), 
                                                         enumImageFileFormat, intImageFileDpi, longImageFileQuality))
                                    {
                                        return false;
                                    }

                                break;
                            }
                        }
                        else
                        {
                            // intAttachmentSeqNo > 0 為 附件
                            // 若為附件, 傳入值做以下加工 Convert.ToInt32("0" + strId_OPDefinition).ToString("00"), 取得 "00" 2碼版本序號
                            switch (enumConversionComponent)
                            {
                                case ConversionComponent.veryPDF_doc2pdf:

                                    if (!Execute_veryPDF_doc2pdf(strdocPrintPathName,
                                                         @"-i " + "\"" + fileInfoDocumentFile.FullName + "\" "
                                                       + @"-o " + "\"" + Path.Combine(strImageFilePath,
                                                                                strWordDocNo + "_" + strDocumentFileName
                                                                                             + "_A" + intAttachmentSeqNo.ToString("00")
                                                                                             + "_" + (strId_OPDefinition == "TBD" ? "TBD" : Convert.ToInt32("0" + strId_OPDefinition).ToString("00"))
                                                                                             + "." + enumImageFileFormat.ToString()) + "\" "
                                                       + @"-d -b 24 -r " + intImageFileDpi.ToString() + " -D -E 2"))
                                    {
                                        return false;
                                    }

                                    break;

                                case ConversionComponent.veryPDF_docPrintCom:

                                    if (!Execute_veryPDF_docPrintCom("", "", fileInfoDocumentFile.FullName,
                                                                Path.Combine(strImageFilePath,
                                                                                strWordDocNo + "_" + strDocumentFileName
                                                                                             + "_A" + intAttachmentSeqNo.ToString("00")
                                                                                             + "_" + (strId_OPDefinition == "TBD" ? "TBD" : Convert.ToInt32("0" + strId_OPDefinition).ToString("00"))
                                                                                             + "." + enumImageFileFormat.ToString()),
                                                                @"-d -b 24 -r " + intImageFileDpi.ToString() + " -D -E 2 -quality " + longImageFileQuality.ToString()))
                                    {
                                        return false;
                                    }

                                    break;
                            
                                case ConversionComponent.easyPDF:

                                    if (!Execute_easyPDF(fileInfoDocumentFile.FullName,
                                                            Path.Combine(strImageFilePath,
                                                                            strWordDocNo + "_" + strDocumentFileName
                                                                                         + "_A" + intAttachmentSeqNo.ToString("00")
                                                                                         + "_" + (strId_OPDefinition == "TBD" ? "TBD" : Convert.ToInt32("0" + strId_OPDefinition).ToString("00"))
                                                                                         + "." + enumImageFileFormat.ToString()),
                                                         enumImageFileFormat, intImageFileDpi, longImageFileQuality))
                                    {
                                        return false;
                                    }

                                    break;
                            }
                        }

                        #endregion

                        break;
                }

                #region 對簽章頁面檔重新命名, 並取得轉換的影像檔資訊

                fileInfoOutputImageFiles
                    = new DirectoryInfo(strImageFilePath).GetFiles(strWordDocNo + "_*." + enumImageFileFormat.ToString());

                foreach (FileInfo fileInfo in fileInfoOutputImageFiles)
                {
                    if (fileInfo.Name.IndexOf("_WS_") > 0)
                    {
                        if (File.Exists(fileInfo.FullName.Replace("_WS_", "_W_").Replace(strId_OPDefinition + "_0",
                                                                             strId_OPDefinition + "_S")))
                        {
                            File.Delete(fileInfo.FullName.Replace("_WS_", "_W_").Replace(strId_OPDefinition + "_0",
                                                                             strId_OPDefinition + "_S"));
                        }

                        fileInfo.MoveTo(fileInfo.FullName.Replace("_WS_", "_W_").Replace(strId_OPDefinition + "_0",
                                                                                            strId_OPDefinition + "_S"));
                    }
                }

                fileInfoOutputImageFiles
                    = new DirectoryInfo(strImageFilePath).GetFiles(strWordDocNo + "_*." + enumImageFileFormat.ToString());

                #endregion

                return true;
            }
            catch (Exception exception)
            {
                strLastErrorMessage = exception.Message;
                fileInfoOutputImageFiles = null;
                return false;
            }
        }

        #region 使用 Dsic.WebOE.Display 元件，將 Di 檔轉換為影像檔

        /// <summary>
        /// 設定 Di 轉換之 XSL 檔絕對路徑及檔名。
        /// </summary>
        private string str_XslCompiledTransformFilePathName;
        /// <summary>
        /// 設定 Di 轉換之 XSL 檔絕對路徑及檔名。
        /// </summary>
        public string strXslCompiledTransformFilePathName
        {
            get
            {
                if (str_XslCompiledTransformFilePathName == null)
                {
                    str_XslCompiledTransformFilePathName = "";
                }

                return str_XslCompiledTransformFilePathName;
            }

            set
            {
                str_XslCompiledTransformFilePathName = value;
            }
        }

        /// <summary>
        /// 將 Di 轉為 XmlDocument。
        /// </summary>
        /// <returns></returns>
        private XmlDocument PreviewDiFileToXmlDocument(string strDiFilePathName)
        {   
            XmlDocument xmlDocument = new XmlDocument();
            MemoryStream memoryStream = new MemoryStream();
            XslCompiledTransform xslCompiledTransform = new XslCompiledTransform();            
            XmlReaderSettings xmlReaderSettings = new XmlReaderSettings();
            xmlReaderSettings.ProhibitDtd = false;
            XmlWriter xmlWriter = null;
            XmlReader xmlReader = null;

            try
            {
                xmlWriter = XmlWriter.Create(memoryStream);
                xmlReader = XmlReader.Create(strDiFilePathName, xmlReaderSettings);

                xslCompiledTransform.Load(strXslCompiledTransformFilePathName, 
                                            new System.Xml.Xsl.XsltSettings(false, true), 
                                            new System.Xml.XmlUrlResolver());
                xslCompiledTransform.Transform(xmlReader, new XsltArgumentList(), xmlWriter);

                memoryStream.Position = 0;
                xmlDocument.Load(memoryStream);
                
                return xmlDocument;
            }
            catch (Exception exception)
            {
                strLastErrorMessage = exception.Message;
                return null;
            }
            finally
            {
                if (xmlReader != null)
                {
                    xmlReader.Close();
                }

                if (xmlWriter != null)
                {
                    xmlWriter.Close();
                }

                memoryStream.Dispose();
            }
        }
    
        /// <summary>
        /// 將 Di 內容之 XmlDocument 輸出為影像檔。
        /// </summary>
        /// <param name="xmlDocument"></param>
        /// <param name="strImageFilePathNameWithoutExtension"></param>
        /// <param name="enumImageFileFormat"></param>
        /// <param name="intImageFileDpi"></param>
        /// <param name="longImageFileQuality"></param>
        /// <returns></returns>
        private bool ExportImageFile(XmlDocument xmlDocument, string strImageFilePathNameWithoutExtension,
                                     ImageFileFormat enumImageFileFormat, int intImageFileDpi, long longImageFileQuality)
        {
            //try
            //{
            //    DocumentReader documentReader = new DocumentReader(xmlDocument.DocumentElement, 0F, 0F);
            //    PictureExporter pictureExporter = new PictureExporter(documentReader.Document);
            //    pictureExporter.Document.Append(new BindLine(pictureExporter.Document,false));
            //    pictureExporter.Document.Append(new PageNumberRule(pictureExporter.Document,false,false,true,"",""));
            //    pictureExporter.Dpi = intImageFileDpi;

            //    // 取得頁數
            //    intPageCount = pictureExporter.Document.TotalPages;
                 
            //    Image[] images = pictureExporter.Paint();
                
            //    EncoderParameter encoderParameter = new EncoderParameter(Encoder.Quality, longImageFileQuality);
            //    EncoderParameters encoderParameters = new EncoderParameters(1);
            //    encoderParameters.Param[0] = encoderParameter;

            //    ImageCodecInfo[] CodecInfo = ImageCodecInfo.GetImageEncoders();
            //    foreach (ImageCodecInfo imageCodecInfo in CodecInfo)
            //    {
            //        switch (enumImageFileFormat)
            //        {
            //            case ImageFileFormat.JPG:

            //                #region JPG

            //                if (imageCodecInfo.MimeType == "image/jpeg")
            //                {
            //                    for (int i = 1; i <= images.Length; i++)
            //                    {
            //                        images[i - 1].Save(strImageFilePathNameWithoutExtension
            //                                            + "_" + i.ToString("0000")
            //                                            + "." + enumImageFileFormat.ToString(),
            //                                            imageCodecInfo, encoderParameters);
            //                    }

            //                    break;
            //                }

            //                #endregion

            //                break;
            //            case ImageFileFormat.TIF:
            //                ImageCodecInfo info = null;
            //                Bitmap pages = null;
            //                foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
            //                    if (ice.FormatID == ImageFormat.Tiff.Guid)
            //                        info = ice;
            //                System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.SaveFlag;
            //                System.Drawing.Imaging.Encoder enc2 = System.Drawing.Imaging.Encoder.Quality;
            //                System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Compression;
            //                System.Drawing.Imaging.Encoder colordepth = System.Drawing.Imaging.Encoder.ColorDepth;

            //                EncoderParameters ep;

            //                ep = new EncoderParameters(4);
            //                ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
            //                ep.Param[1] = new EncoderParameter(enc2, 100L);
            //                ep.Param[2] = new EncoderParameter(myEncoder, (long)EncoderValue.CompressionLZW);
            //                ep.Param[3] = new EncoderParameter(colordepth, 24L);

            //                for (int i = 0; i < images.Length; i++)
            //                {
            //                    if (i == 0)
            //                    {
            //                        pages = (Bitmap)images[i];
            //                        pages.Save(strImageFilePathNameWithoutExtension + ".tif", info, ep);
            //                    }
            //                    else
            //                    {
            //                        ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);
            //                        Bitmap bm = (Bitmap)images[i];
            //                        pages.SaveAdd(bm, ep);
            //                    }
            //                }
            //                ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);
            //                pages.SaveAdd(ep);
            //                ep.Dispose(); 
            //                break;

            //            default:
            //                break;
            //        }
            //    }

            //    encoderParameter.Dispose();
            //    encoderParameters.Dispose();

                return true;
            //}
            //catch (Exception exception)
            //{
            //    strLastErrorMessage = exception.Message;
            //    return false;
            //}
        }

        #endregion

        #region 使用 docPrint 元件，將支援 printable 的文件檔轉換為影像檔

        /// <summary>
        /// 設定 docPrint 元件執行檔絕對路徑及檔名。
        /// </summary>
        private string str_docPrintPathName;
        /// <summary>
        /// 設定 docPrint 元件執行檔絕對路徑及檔名。
        /// </summary>
        public string strdocPrintPathName
        {
            get
            {
                if (str_docPrintPathName == null)
                {
                    str_docPrintPathName = @"C:\Program Files\docPrint Pro v4.0\doc2pdf.exe";
                }

                return str_docPrintPathName;
            }

            set
            {
                str_docPrintPathName = value;
            }
        }

        /// <summary>
        /// 設定 docPrint 元件執行 Timeout 秒數。
        /// </summary>
        /// <remarks>若未設定，則預計允許 60秒。</remarks>
        private int int_docPrintTimeout = 60;
        /// <summary>
        /// 設定 docPrint 元件執行 Timeout 秒數。
        /// </summary>
        /// <remarks>若未設定，則預計允許 60秒。</remarks>
        public int intdocPrintTimeout
        {
            get
            {
                return int_docPrintTimeout;
            }

            set
            {
                int_docPrintTimeout = value;
            }
        }

        /// <summary>
        /// 使用 docPrint 元件，執行 Command Line (doc2pdf.exe)，轉數位內容頁面檔。
        /// </summary>
        /// <param name="strStartInfoFileName"></param>
        /// <param name="strStartInfoArguments"></param>
        /// <returns></returns>
        private bool Execute_veryPDF_doc2pdf(string strStartInfoFileName, string strStartInfoArguments)
        {
            Process process = new Process();
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.FileName = strStartInfoFileName;
            process.StartInfo.Arguments = strStartInfoArguments;

            try
            {
                process.Start();

                while (!process.HasExited)
                {
                    #region 若超過設定秒數，則取消

                    if (process.TotalProcessorTime > new TimeSpan(0, 0, intdocPrintTimeout))
                    {
                        process.Close();

                        strLastErrorMessage = "[Timeout] 執行時間超過 " + intdocPrintTimeout.ToString() + "秒。";
                        return false;
                    }

                    #endregion
                }

                process.Close();

                return true;
            }
            catch (Exception exception)
            {
                strLastErrorMessage = exception.Message;
                return false;
            }
            finally
            {
                process.Dispose();
            }
        }

        /// <summary>
        /// 使用 docPrint 元件，執行 SDK (docPrintCom.dll)，轉數位內容頁面檔。
        /// </summary>
        /// <param name="strUserName"></param>
        /// <param name="strPassword"></param>
        /// <param name="strInputFile"></param>
        /// <param name="strOutputFile"></param>
        /// <param name="strOptions"></param>
        /// <returns></returns>
        private bool Execute_veryPDF_docPrintCom(string strUserName, string strPassword, string strInputFile, string strOutputFile, string strOptions)
        {
            DOCPRINTCOMLib.docPrintClass clsdocPrintClass = new DOCPRINTCOMLib.docPrintClass();
            clsdocPrintClass.docPrintCOM_Register("VL6IEN2GRW8EYN8P1D7F", "Digitware System Integration Corporation");

            try
            {
                clsdocPrintClass.docPrintCOM_DocumentConverterEx(strUserName, strPassword, strInputFile, strOutputFile, strOptions);

                if (clsdocPrintClass.docPrintCOM_GetLastError > 1)
                {
                    strLastErrorMessage = "[docPrintCOM Error Code: " + clsdocPrintClass.docPrintCOM_GetLastError.ToString() + " ]";
                    return false;
                }
                
                return true;
            }
            catch (Exception exception)
            {
                strLastErrorMessage = exception.Message;
                return false;
            }
        }


        public bool do_veryPDF_docPrintCom(string _strInputFile, string _strOutputFile, string _strOptions)
        {
            if (_strOptions=="")
                return Execute_veryPDF_docPrintCom("", "", _strInputFile, _strOutputFile, @"-d -b 24 -r 300 -D -E 2"); 
            else
                return Execute_veryPDF_docPrintCom("", "", _strInputFile, _strOutputFile, _strOptions);
        }
        #endregion

        #region 使用 BCL EasyPDF Printer 元件轉換為影像檔

        /// <summary>
        /// 轉出影像檔頁數。
        /// </summary>
        private int int_PageCount;
        /// <summary>
        /// 轉出影像檔頁數。
        /// </summary>
        public int intPageCount
        {
            get
            {
                return int_PageCount;
            }

            set
            {
                int_PageCount = value;
            }
        }

        /// <summary>
        /// 若頁數大於設定數量，終止轉檔。
        /// </summary>
        /// <remarks>若未設定，則預設值為10頁。</remarks>
        private int int_PageCountConstrain = 10;
        /// <summary>
        /// 若頁數大於設定數量，終止轉檔。
        /// </summary>
        /// <remarks>若未設定，則預設值為10頁。</remarks>
        public int intPageCountConstrain
        {
            get
            {
                return int_PageCountConstrain;
            }

            set
            {
                int_PageCountConstrain = value;
            }
        }

        /// <summary>
        /// 設定轉檔 Timeout 值。
        /// </summary>
        /// <remarks>若未設定，則預設值為60000毫秒。</remarks>
        private int int_FileConversionTimeout = 60 * 1000;
        /// <summary>
        /// 設定轉檔 Timeout 值。
        /// </summary>
        /// <remarks>若未設定，則預設值為60000毫秒。</remarks>
        public int intFileConversionTimeout
        {
            get
            {
                return int_FileConversionTimeout;
            }

            set
            {
                int_FileConversionTimeout = value;
            }
        }

        private bool Execute_easyPDF(
            string strInputFilePathName, string strOutputFilePathName,
            ImageFileFormat enumImageFileFormat, int intImageFileDpi, long longImageFileQuality)
        {
            try
            {
                #region 不使用 easyPDF 元件, 故 remark

                //FileInfo fileInfoInputFile = new FileInfo(strInputFilePathName);
                //FileInfo fileInfoOutputFile = new FileInfo(strOutputFilePathName);

                //#region 將來源檔轉為 PDF

                //Type typeEasyPDFPrinter = Type.GetTypeFromProgID("easyPDF.Printer.6");
                
                //Printer oPrinter = (Printer)Activator.CreateInstance(typeEasyPDFPrinter);
                                
                //// 選用 來源格式類別
                //switch (oPrinter.GetPrintJobTypeOf(fileInfoInputFile.FullName))
                //{
                //    #region MS Office

                //    case prnPrintJobType.PRN_PRNJOB_WORD:
                //        WordPrintJob printJobWord = oPrinter.WordPrintJob;

                //        try
                //        {
                //            printJobWord.PrintInBackground = true;
                //            printJobWord.ConvertHyperlinks = false;
                //            //printJobWord.ConvertBookmarks = false;

                //            // Timeout 設定，單位為毫秒
                //            printJobWord.FileConversionTimeout = int_FileConversionTimeout;

                //            // 設定字集參數
                //            printJobWord.PDFSetting.FontEmbedAsType0 = true;
                //            printJobWord.PDFSetting.FontEmbedding = prnFontEmbedding.PRN_FONT_EMBED_SUBSET;
                //            printJobWord.PDFSetting.FontSubstitution = prnFontSubstitution.PRN_FONT_SUBST_PDF;

                //            printJobWord.PrintOut(
                //                fileInfoInputFile.FullName,
                //                fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //        }
                //        catch
                //        {
                //            strLastErrorMessage = printJobWord.PrinterResultMessage;
                //            return false;
                //        }

                //        break;

                //    case prnPrintJobType.PRN_PRNJOB_EXCEL:
                //        ExcelPrintJob printJobExcel = oPrinter.ExcelPrintJob;

                //        try
                //        {
                //            printJobExcel.PrintAllSheets = true;
                //            // Timeout 設定，單位為毫秒
                //            printJobExcel.FileConversionTimeout = int_FileConversionTimeout;
                //            printJobExcel.FitToPagesWide = 1;

                //            // 設定字集參數
                //            printJobExcel.PDFSetting.FontEmbedAsType0 = true;
                //            printJobExcel.PDFSetting.FontEmbedding = prnFontEmbedding.PRN_FONT_EMBED_SUBSET;
                //            printJobExcel.PDFSetting.FontSubstitution = prnFontSubstitution.PRN_FONT_SUBST_PDF;

                //            printJobExcel.PrintOut(
                //                fileInfoInputFile.FullName,
                //                fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //        }
                //        catch
                //        {
                //            strLastErrorMessage = printJobExcel.PrinterResultMessage;
                //            return false;
                //        }

                //        break;

                //    case prnPrintJobType.PRN_PRNJOB_PPT:
                //        PowerPointPrintJob printJobPowerPoint = oPrinter.PowerPointPrintJob;

                //        try
                //        {
                //            // Timeout 設定，單位為毫秒
                //            printJobPowerPoint.FileConversionTimeout = int_FileConversionTimeout;

                //            // 設定字集參數
                //            printJobPowerPoint.PDFSetting.FontEmbedAsType0 = true;
                //            printJobPowerPoint.PDFSetting.FontEmbedding = prnFontEmbedding.PRN_FONT_EMBED_SUBSET;
                //            printJobPowerPoint.PDFSetting.FontSubstitution = prnFontSubstitution.PRN_FONT_SUBST_PDF;

                //            printJobPowerPoint.PrintOut(
                //                fileInfoInputFile.FullName,
                //                fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //        }
                //        catch
                //        {
                //            strLastErrorMessage = printJobPowerPoint.PrinterResultMessage;
                //            return false;
                //        }

                //        break;

                //    #endregion

                //    #region IE ( 會印出 N of M 頁 及 URL, 故不使用, .htm 類型的檔案強制以 IEExtended 物件處理 )

                //    //case prnPrintJobType.PRN_PRNJOB_IE:
                //    //    IEPrintJob printJobIE = oPrinter.IEPrintJob;

                //    //    try
                //    //    {
                //    //        // Timeout 設定，單位為毫秒
                //    //        printJobIE.FileConversionTimeout = int_FileConversionTimeout;

                //    //        printJobIE.ResumeOnError = true;
                //    //        printJobIE.IESetting.DisableScriptDebugger = true;
                //    //        printJobIE.IESetting.DisplayErrorDialogOnEveryError = false;

                //    //        // 設定字集參數
                //    //        printJobIE.PDFSetting.FontEmbedAsType0 = true;
                //    //        printJobIE.PDFSetting.FontEmbedding = prnFontEmbedding.PRN_FONT_EMBED_SUBSET;
                //    //        printJobIE.PDFSetting.FontSubstitution = prnFontSubstitution.PRN_FONT_SUBST_PDF;

                //    //        printJobIE.PrintOut(
                //    //            fileInfoInputFile.FullName,
                //    //            fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //    //    }
                //    //    catch
                //    //    {
                //    //        strLastErrorMessage = printJobIE.PrinterResultMessage;
                //    //        return false;
                //    //    }

                //    //    break;

                //    #endregion

                //    #region IEExtended

                //    case prnPrintJobType.PRN_PRNJOB_IE:
                //        IEExtendedPrintJob printJobIEExtended = oPrinter.IEExtendedPrintJob;

                //        try
                //        {
                //            // Timeout 設定，單位為毫秒
                //            printJobIEExtended.FileConversionTimeout = int_FileConversionTimeout;

                //            printJobIEExtended.ResumeOnError = true;
                //            printJobIEExtended.IEExtendedSetting.DisableScriptDebugger = true;
                //            printJobIEExtended.IEExtendedSetting.DisplayErrorDialogOnEveryError = false;

                //            // 設定字集參數
                //            printJobIEExtended.PDFSetting.FontEmbedAsType0 = true;
                //            printJobIEExtended.PDFSetting.FontEmbedding = prnFontEmbedding.PRN_FONT_EMBED_SUBSET;
                //            printJobIEExtended.PDFSetting.FontSubstitution = prnFontSubstitution.PRN_FONT_SUBST_PDF;

                //            printJobIEExtended.PrintOut(
                //                fileInfoInputFile.FullName,
                //                fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //        }
                //        catch
                //        {
                //            strLastErrorMessage = printJobIEExtended.PrinterResultMessage;
                //            return false;
                //        }

                //        break;

                //    #endregion

                //    #region Image

                //    case prnPrintJobType.PRN_PRNJOB_IMAGE:
                //        IImagePrintJob printJobIImage = oPrinter.ImagePrintJob;

                //        try
                //        {
                //            // Timeout 設定，單位為毫秒
                //            printJobIImage.FileConversionTimeout = int_FileConversionTimeout;

                //            printJobIImage.PrintOut(
                //                fileInfoInputFile.FullName,
                //                fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //        }
                //        catch
                //        {
                //            strLastErrorMessage = printJobIImage.PrinterResultMessage;
                //            return false;
                //        }

                //        break;

                //    #endregion

                //    #region Others

                //    case prnPrintJobType.PRN_PRNJOB_GENERIC:
                //        GenericPrintJob printJobGeneric = oPrinter.GenericPrintJob;
                //        try
                //        {
                //            // Timeout 設定，單位為毫秒
                //            printJobGeneric.FileConversionTimeout = int_FileConversionTimeout;

                //            // 設定字集參數
                //            printJobGeneric.PDFSetting.FontEmbedAsType0 = true;
                //            printJobGeneric.PDFSetting.FontEmbedding = prnFontEmbedding.PRN_FONT_EMBED_SUBSET;
                //            printJobGeneric.PDFSetting.FontSubstitution = prnFontSubstitution.PRN_FONT_SUBST_PDF;

                //            printJobGeneric.PrintOut(
                //                fileInfoInputFile.FullName,
                //                fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));
                //        }
                //        catch
                //        {
                //            strLastErrorMessage = printJobGeneric.PrinterResultMessage;
                //            return false;
                //        }

                //        break;

                //    #endregion

                //    default:
                //        strLastErrorMessage = "無法將 " + fileInfoInputFile.FullName + " 轉換為 數位內容頁面檔！";
                //        break;
                //}

                //#endregion

                //#region 取得檔案頁數

                //Type typeEasyPDFProcessor = Type.GetTypeFromProgID("easyPDF.PDFProcessor.6");
                //IPDFProcessor oProcessor = (IPDFProcessor)Activator.CreateInstance(typeEasyPDFProcessor);

                //intPageCount
                //    = oProcessor.GetPageCount(fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"), null);

                //#endregion

                //// 若頁數小於設定數量才予轉檔，否則終止
                //if (intPageCount <= intPageCountConstrain)
                //{
                //    #region 將 PDF 轉為 影像檔

                //    Type typeEasyPDFConverter = Type.GetTypeFromProgID("easyPDF.PDFConverter.6");
                //    IPDFConverter oConverter = (IPDFConverter)Activator.CreateInstance(typeEasyPDFConverter);

                //    PDF2Image oPDF2Image = oConverter.PDF2Image;

                //    // DPI 設定
                //    oPDF2Image.ImageResolution = intImageFileDpi;
                //    oPDF2Image.ImageColor = cnvImageColor.CNV_IMAGE_CLR_24BIT;

                //    // 轉出格式設定
                //    switch (enumImageFileFormat)
                //    {
                //        case ImageFileFormat.TIF:
                //            oPDF2Image.ImageFormat = cnvImageFormat.CNV_IMAGE_FMT_TIFF;
                //            break;

                //        default:
                //            oPDF2Image.ImageFormat = cnvImageFormat.CNV_IMAGE_FMT_JPEG;
                //            break;
                //    }

                //    // Quality 設定
                //    oPDF2Image.ImageQuality = Convert.ToInt32(longImageFileQuality);
                //    // Timeout 設定，單位為毫秒
                //    oPDF2Image.FileConversionTimeout = int_FileConversionTimeout;
                //    // 頁碼長度設定
                //    oPDF2Image.MinimumPageNumberDigits = 4;
                //    // 分隔符號
                //    oPDF2Image.PageNumberSeparator = "_";
                //    // 只有一頁的檔案也加上 _0001 頁碼
                //    oPDF2Image.AddPageNumberForSingleFileOutput = true;
                //    // 由 PDF 轉換為影像檔
                //    oPDF2Image.Convert(fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"),
                //                       fileInfoOutputFile.FullName.Replace(fileInfoOutputFile.Extension.ToUpper(),
                //                                                           fileInfoOutputFile.Extension.ToLower()), 
                //                       null, 1, intPageCount);
                    
                //    // 如果 _0001_0001 檔案存在
                //    if (File.Exists(fileInfoOutputFile.FullName.Replace(fileInfoOutputFile.Extension, "_0001_0001" + fileInfoOutputFile.Extension)))
                //    {
                //        // 如果 _0001 檔案存在
                //        if (File.Exists(fileInfoOutputFile.FullName.Replace(fileInfoOutputFile.Extension, "_0001" + fileInfoOutputFile.Extension)))
                //        {
                //            File.Delete(fileInfoOutputFile.FullName.Replace(fileInfoOutputFile.Extension, "_0001" + fileInfoOutputFile.Extension));
                //        }

                //        File.Move(fileInfoOutputFile.FullName.Replace(fileInfoOutputFile.Extension, "_0001_0001" + fileInfoOutputFile.Extension),
                //                  fileInfoOutputFile.FullName.Replace(fileInfoOutputFile.Extension, "_0001" + fileInfoOutputFile.Extension));
                //    }

                //    #endregion
                //}

                //// 刪除暫存之 PDF 檔
                //File.Delete(fileInfoInputFile.FullName.Replace(fileInfoInputFile.Extension, ".PDF"));

                #endregion

                return true;
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                strLastErrorMessage = exception.Message;
                return false;
            }
            catch (Exception exception)
            {
                strLastErrorMessage = exception.Message;
                return false;
            }
        }

        #endregion
    }
}