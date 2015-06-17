using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using DigitalSealed;
using DigitalSealed.Tools;
using System.IO;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo;
using DigitalSealed.SignOnline;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object;
using DigitalSealed.OA;

namespace testDigitalSealed
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Signature sig = new Signature();
            sig.CardReaderSlot = 0; //選擇使用第幾台讀卡機
            string start = sig.GetPGKIStartEndDate(0);
            string end = sig.GetPGKIStartEndDate(1);
        }

        //private bool Download(string _docno, string localFilePathName)
        //{
        //    bool bRes = false;
        //    int ChunkSize = 0;
        //    int MaxRetries = 6;
        //    int c_ChunkSize = 16 * 1024;// max number of corrupted chunks or failed transfers to allow before giving up
        //    long MaxRequestLength = 4096;	// default, this is updated so that the transfer class knows how much the server will accept
        //    int PreferredTransferDuration = 800;		// milliseconds, the timespan the class will attempt to achieve for each chunk, to give responsive feedback on the progress bar.
        //    int ChunkSizeSampleInterval = 15;
        //    int NumRetries = 0;
        //    int numIterations = 0;	
        //    string localFilePath, localFileName, remoteFilePath, remoteFileName;
        //    long Offset = 0;

        //    ChunkSize = c_ChunkSize;
            
        //    DCPWS.DCPWSE svcClient = new DCPWS.DCPWSE(); 

        //    localFileName = System.IO.Path.GetFileName(localFilePathName);
        //    localFilePath = localFilePathName.Substring(0, localFilePathName.Length - localFileName.Length);

        //    if (0 == localFilePath.Length)
        //    {
        //        bRes = false;
        //        goto Exit;
        //    }

        //    if (0 == localFileName.Length)
        //    {
        //        localFileName = _docno + ".dsicdp";
        //    }
        //    localFilePathName = Path.Combine(localFilePath, localFileName);
        //    try
        //    {
        //        long FileSize = svcClient.GetFileSize(_docno);
        //        if (0 >= FileSize)
        //        {
        //            bRes = false;
        //            goto Exit;
        //        }

        //        ChunkSize = c_ChunkSize;
        //        Offset = 0;
        //        if (System.IO.File.Exists(localFilePathName))   // create a new empty file
        //            System.IO.File.Create(localFilePathName).Close();

        //        DateTime StartTime = DateTime.Now;
        //        using (FileStream fs = new FileStream(localFilePathName, FileMode.OpenOrCreate, FileAccess.Write))
        //        {
        //            fs.Seek(Offset, SeekOrigin.Begin);

        //            while (Offset < FileSize)
        //            {
        //                int currentIntervalMod = numIterations % ChunkSizeSampleInterval;
        //                if (currentIntervalMod == 0)
        //                    StartTime = DateTime.Now;	// used to calculate the time taken to transfer the first 5 chunks
        //                else if (currentIntervalMod == 1)
        //                {
        //                    double transferTime = DateTime.Now.Subtract(StartTime).TotalMilliseconds;
        //                    double averageBytesPerMilliSec = ChunkSize / transferTime;
        //                    double preferredChunkSize = averageBytesPerMilliSec * PreferredTransferDuration;
        //                    ChunkSize = (int)Math.Min(MaxRequestLength, Math.Max(4 * 1024, preferredChunkSize));
        //                }
        //                try
        //                {
        //                    byte[] Buffer = svcClient.DownloadChunk(_docno, Offset, ChunkSize);
        //                    fs.Write(Buffer, 0, Buffer.Length);
        //                    Offset += Buffer.Length;	// save the offset position for resume
        //                }
        //                catch (Exception ex)
        //                {
        //                    if (NumRetries++ >= MaxRetries)	// too many retries, bail out
        //                    {
        //                        bRes = false;
        //                        goto Exit;
        //                    }
        //                }
        //                numIterations++;
        //            }
        //            try
        //            {
        //                fs.Close();
        //            }
        //            catch (Exception ex)
        //            {
        //                bRes = false;
        //            }
        //        }
        //        bRes = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        bRes = false;
        //        goto Exit;
        //    }

        //Exit:
        //    return bRes;
        //}

        //private bool uploadfile(string _docno)
        //{
        //    int ChunkSize = 0;
            
        //    bool bRes = false;
        //    int c_ChunkSize = 16 * 1024;	// 16Kb default
        //    int MaxRetries = 6;						
        //    long MaxRequestLength = 4096;	
        //    int PreferredTransferDuration = 800;	
        //    int ChunkSizeSampleInterval = 15;
        //    long Offset = 0;
        //    ChunkSize = c_ChunkSize;
        //    byte[] Buffer = new byte[ChunkSize];
        //    using (FileStream fs = new FileStream("d:\\temp\\" + _docno + ".dsicdp", FileMode.Open, FileAccess.Read))
        //    {
        //        fs.Position = Offset;
        //        int BytesRead;
        //        int NumRetries = 0;
        //        int numIterations = 0;
        //        DateTime StartTime = DateTime.Now;
        //        do
        //        {
        //            BytesRead = fs.Read(Buffer, 0, ChunkSize);	// read the next chunk (if it exists) into the buffer.  the while loop will terminate if there is nothing left to read

        //            // check if this is the last chunk and resize the bufer as needed to avoid sending a mostly empty buffer (could be 10Mb of 000000000000s in a large chunk)
        //            if (BytesRead != Buffer.Length)
        //            {
        //                ChunkSize = BytesRead;
        //                byte[] TrimmedBuffer = new byte[BytesRead];
        //                Array.Copy(Buffer, TrimmedBuffer, BytesRead);
        //                Buffer = TrimmedBuffer;
        //            }
        //            if (Buffer.Length == 0)
        //                break;
        //            try
        //            {
        //                DCPWS.DCPWSE svcClient = new DCPWS.DCPWSE();
        //                svcClient.UploadChunk(_docno, Buffer, Offset);
                        
        //                int currentIntervalMod = numIterations % ChunkSizeSampleInterval;
        //                if (currentIntervalMod == 0)	// start the timer for this chunk
        //                    StartTime = DateTime.Now;
        //                else if (currentIntervalMod == 1)
        //                {
        //                    double transferTime = DateTime.Now.Subtract(StartTime).TotalMilliseconds;
        //                    double averageBytesPerMilliSec = ChunkSize / transferTime;
        //                    double preferredChunkSize = averageBytesPerMilliSec * PreferredTransferDuration;
        //                    ChunkSize = (int)Math.Min(MaxRequestLength, Math.Max(4 * 1024, preferredChunkSize));	// the sample chunk has been transferred, calc the time taken and adjust the chunk size accordingly
        //                    Buffer = new byte[ChunkSize];    // reset the size of the buffer for the new chunksize
        //                }

        //                // Offset is only updated AFTER a successful send of the bytes. this helps the 'retry' feature to resume the upload from the current Offset position if AppendChunk fails.
        //                Offset += BytesRead;	// save the offset position for resume
        //            }
        //            catch (Exception ex)
        //            {
        //                fs.Position -= BytesRead;

        //                if (NumRetries++ >= MaxRetries) // too many retries, bail out
        //                {
        //                    bRes = false;
        //                    goto Exit;
        //                }
        //            }
        //            numIterations++;
        //        } while (BytesRead > 0);
        //        try
        //        {
        //            fs.Close();
        //            bRes = true;
        //        }
        //        catch (Exception ex)
        //        {
        //            bRes = false;
        //        }
        //    }
        //Exit:
        //    return bRes;
        //}

        private void fileCompress(MakeDCP mdcp)
        {
            //將封裝檔存檔並刪除暫存檔
            //mdcp.封裝檔XML.Save(Path.Combine(mdcp.檔案來源資料夾, mdcp.公文文號 + ".si"));

            ////將所有封裝內容做壓縮
            //if (!mdcp.壓縮封裝檔(Path.Combine(mdcp.檔案來源資料夾, mdcp.公文文號 + ".si"), @"D:\temp"))
            //    MessageBox.Show("壓縮封裝檔失敗!");
            //else
            //{
            //    if (uploadfile(mdcp.公文文號))
            //        MessageBox.Show("封裝成功!");
            //    else
            //        MessageBox.Show("上傳失敗!");
            //}
        }

        private void 承辦_Click(object sender, EventArgs e)
        {
            //傳送者
            ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "14", "專員", "Jack", "Jack");
            
            //接收者Option
            ModifyInfo.signer re = new ModifyInfo.signer("秘書室", "14", "科長", "DSIC", "DSIC");
            
            //建立封裝類別
            MakeDCP mdcp = new MakeDCP("1000000016", sr, null, "承辦意見", FlowInfo.FlowType.呈核);
            
            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";
            
            //所有要加進封裝的檔案存放路徑
            mdcp.檔案來源資料夾 = @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\00016";
            
            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            //att.Add("來文附件1.txt");
            //att.Add("來文附件2.docx");
            string rtnmsg = mdcp.產生來文文件夾("099100000", "函", "200707020001.di", att);
            //使用WebService取來文資料
            //DCPWS.DCPWSE ws =new DCPWS.DCPWSE();
            //string rtnmsg = mdcp.產生來文文件夾(ws.getInDocAttInfo("1", "0990000470")); 
            att.Clear();
            att.Add("公司OPV1.txt");
            //紙本來文
            //att.Add("來文附件.tif");
            //att.Add("2460200A00_ATTCH27.doc"); 
            //string rtnmsg = mdcp.產生來文文件夾("099100000", "函", "200707020001.tif", att);

            rtnmsg += mdcp.加入文書封裝("便簽", "簽", "100JP00016V1.di", att);

            //更新si封裝描述檔,不轉頁面
            if (rtnmsg == "")
            {
                rtnmsg = mdcp.異動封裝檔("serial:59b0b931000100000b3a", false);
            }
            if (rtnmsg != "")
                MessageBox.Show(rtnmsg);
            else
            {
                //將簽核意見htm複製到來資料夾一起壓縮
                File.Copy( @"C:\Program Files\DigitWare\NCC-WEBOE\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾,"100JP00016DcpSignInfo.htm"),true); 
                //將所有封裝內容做壓縮
                fileCompress(mdcp);
            }
        }

        private void 核稿_Click(object sender, EventArgs e)
        {
            //傳送者
            ModifyInfo.signer send = new ModifyInfo.signer("秘書室", "14", "科長", "DSIC", "DSIC");

            //接收者
            ModifyInfo.signer rcv = new ModifyInfo.signer("秘書室", "14", "主任", "Joe", "Joe");

            //上一流程所產生的封裝壓縮檔位置,來源為fileserver上的文號資料夾
            FileInfo finfo = new FileInfo(@"C:\Users\jack\Desktop\1020801393.dsicdp");

            //解壓縮上一簽核點所有檔的存放資料夾,通常等於目前簽核點的來源資料夾
            DirectoryInfo dinfo = new DirectoryInfo(@"C:\Users\jack\Desktop\unzip\00006");

            //建立封裝類別
            MakeDCP mdcp = new MakeDCP(finfo, dinfo, send, rcv, "同意", FlowInfo.FlowType.呈核);

            mdcp.產生來文參考(); 
            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";

            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            //att.Add("來文附件1.txt");
            //att.Add("來文附件2.docx");
            //string rtnmsg = mdcp.產生來文文件夾("099100000", "函", "200707020001.di", att);
            //針對有異動的附件檔案檔名增加版本並加入封裝
            File.Copy(Path.Combine(mdcp.檔案來源資料夾, "公司OPV1.txt"), Path.Combine(mdcp.檔案來源資料夾, "公司OPV2.txt"),true);
            att.Add("公司OPV2.txt");
            //文書本文的部分一律增加版次,先用copy產生假資料
            File.Copy(Path.Combine(mdcp.檔案來源資料夾, "100JP00016V1.di"), Path.Combine(mdcp.檔案來源資料夾, "100JP00016V2.di"),true);
            string rtn = mdcp.加入文書封裝("便簽", "簽", "100JP00016V2.di", att);
            //先用copy產生假資料
            File.Copy(Path.Combine(mdcp.檔案來源資料夾, "100JP00016V1.di"), Path.Combine(mdcp.檔案來源資料夾, "100RM00006V1.di"),true);
            rtn += mdcp.加入文書封裝("開會通知單", "函", "100RM00006V1.di", null);

            //更新si封裝描述檔,不轉頁面
            rtn += mdcp.異動封裝檔("12345678", false);
            if (rtn != "")
            {
                MessageBox.Show(rtn);
                return;
            }
            //將簽核意見htm複製到來源資料夾一起壓縮
            File.Copy(@"C:\Program Files\DigitWare\NCC-WEBOE\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾, "100RM00006DcpSignInfo.htm"),true); 
            //將所有封裝內容做壓縮
            fileCompress(mdcp);
        }

        private void 決行_Click(object sender, EventArgs e)
        {
            //傳送者
            ModifyInfo.signer send = new ModifyInfo.signer("秘書室", "14", "主任", "Joe", "Joe");

            //上一流程所產生的封裝壓縮檔位置,來源為fileserver上的文號資料夾
            FileInfo finfo = new FileInfo(@"D:\temp\1000000016.dsicdp");

            //解壓縮上一簽核點所有檔的存放資料夾,通常等於目前簽核點的來源資料夾
            DirectoryInfo dinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\00111");

            //建立封裝類別,接收者可為null
            MakeDCP mdcp = new MakeDCP(finfo, dinfo, send, null, "決", FlowInfo.FlowType.決行);

            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";

            //準備要加入封裝的檔案
            List<string> att = new List<string>();

            //文書本文的部分一律增加版次,先用copy產生假資料
            File.Copy(Path.Combine(mdcp.檔案來源資料夾, "100JP00016V1.di"), Path.Combine(mdcp.檔案來源資料夾, "100JP00016V3.di"),true);
            string rtn = mdcp.加入文書封裝("便簽", "簽", "100JP00016V3.di", att);
            File.Copy(Path.Combine(mdcp.檔案來源資料夾, "100RM00006V1.di"), Path.Combine(mdcp.檔案來源資料夾, "100RM00006V2.di"),true);
            rtn += mdcp.加入文書封裝("開會通知單", "函", "100RM00006V2.di", null);

            //將文稿htm檔複製到來源資料夾供轉頁面
            File.Copy(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\100JP00016V3.htm", Path.Combine(mdcp.檔案來源資料夾, "100JP00016V3.htm"), true);
            File.Copy(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\100RM00006V2.htm", Path.Combine(mdcp.檔案來源資料夾, "100RM00006V2.htm"), true);
            //將簽核資訊htm檔(DcpSignInfo.htm)複製到來源資料夾 : 文書號DcpSignInfo.htm
            File.Copy(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾, "100JP00016DcpSignInfo.htm"), true);
            File.Copy(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾, "100RM00006DcpSignInfo.htm"), true);
            //更新si封裝描述檔,並轉頁面
            rtn += mdcp.異動封裝檔("731022", true);
            if (rtn != "")
            {
                MessageBox.Show(rtn);
                return;
            }

            
            //將所有封裝內容做壓縮
            fileCompress(mdcp);
        }

        private void 並會一_Click(object sender, EventArgs e)
        {
            //傳送者
            ModifyInfo.signer send = new ModifyInfo.signer("資訊室", "34", "會辦人", "Jay", "Jay");
            //上一流程所產生的封裝壓縮檔位置,來源為fileserver上的文號資料夾
            FileInfo finfo = new FileInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\1000000016.dsicdp");

            //解壓縮上一簽核點所有檔的存放資料夾,通常等於目前簽核點的來源資料夾
            DirectoryInfo dinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\split01");

            MakeDCP mdcp = new MakeDCP(finfo, dinfo, send, null, "會辦意見OK", FlowInfo.FlowType.並會);
            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";
            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            string rtn = mdcp.加入文書封裝("便簽", "簽", "100JP00016V3.di", att);
            rtn += mdcp.加入文書封裝("開會通知單", "函", "100RM00006V2.di", null);

            rtn += mdcp.異動封裝檔("731022", false);
            if (rtn != "")
            {
                MessageBox.Show(rtn);
                return;
            }
            //將封裝檔存檔並刪除暫存檔,並會不壓縮
            mdcp.封裝檔XML.Save(Path.Combine(mdcp.檔案來源資料夾, "100JP00016.si"));
            MessageBox.Show("OK");
        }

        private void 並會二_Click(object sender, EventArgs e)
        {
            //傳送者
            ModifyInfo.signer send = new ModifyInfo.signer("工程室", "66", "會辦人", "Jimmy", "Jimmy");
            //上一流程所產生的封裝壓縮檔位置,來源為fileserver上的文號資料夾
            FileInfo finfo = new FileInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\1000000016.dsicdp");

            //解壓縮上一簽核點所有檔的存放資料夾,通常等於目前簽核點的來源資料夾
            DirectoryInfo dinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\split02");

            MakeDCP mdcp = new MakeDCP(finfo, dinfo, send, null, "會辦意見", FlowInfo.FlowType.並會);
            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";
            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            string rtn = mdcp.加入文書封裝("便簽", "簽", "100JP00016V3.di", att);
            rtn += mdcp.加入文書封裝("開會通知單", "函", "100RM00006V2.di", null);
            rtn += mdcp.異動封裝檔("731022", false);
            if (rtn != "")
            {
                MessageBox.Show(rtn);
                return;
            }
            //將封裝檔存檔並刪除暫存檔,並會不壓縮
            mdcp.封裝檔XML.Save(Path.Combine(mdcp.檔案來源資料夾, "100JP00016.si"));
            MessageBox.Show("OK");
        }

        private void 合併_Click(object sender, EventArgs e)
        {
            //回承辦人做合併並會的簽核點定義資料
            ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "14", "專員", "Jack", "Jack");
            //並會前所產生的封裝壓縮檔位置,來源為fileserver上的文號資料夾
            FileInfo finfo = new FileInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\temp\1000000016.dsicdp");

            //解壓縮上一簽核點所有檔的存放資料夾,通常等於目前簽核點的來源資料夾
            DirectoryInfo dinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\merge");

            //建立封裝類別
            MakeDCP mdcp = new MakeDCP(finfo, dinfo, sr, null, "同意", FlowInfo.FlowType.呈核);

            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";

            FileInfo[] depsi = new FileInfo[2];
            depsi[0] = new FileInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\split01\100JP00016.si");
            depsi[1] = new FileInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\split02\100JP00016.si");
            mdcp.合併並會資料(depsi);

            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            string rtn = mdcp.加入文書封裝("便簽", "簽", "100JP00016V3.di", att);
            rtn += mdcp.加入文書封裝("開會通知單", "函", "100RM00006V2.di", null);

            rtn += mdcp.異動封裝檔("731022", false);
            if (rtn != "")
            {
                MessageBox.Show(rtn);
                return;
            }
            //將所有封裝內容做壓縮
            fileCompress(mdcp);
        }

        private void 併案_Click(object sender, EventArgs e)
        {
            string[] childs = new string[2]{"09655588885","09655588883"};
            /////////////先處理子文09655588883/////////////////
            ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "14", "專員", "Jack", "Jack");
            MakeDCP dcp1 = new MakeDCP("09655588883", sr, null, "3份來文彙辦處理", FlowInfo.FlowType.彙併辦);
            dcp1.檔案來源資料夾 = @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\88883";
            List<string> att = new List<string>();
            att.Add("來文附件0970000023.doc");
            dcp1.產生來文文件夾("0970000023", "函", "來文0970000023.di", att);
            //利用webservice找出彙併辦
            dcp1.設定彙併辦("09655588882", childs);
            dcp1.異動封裝檔("731022", false);
            dcp1.封裝檔XML.Save(Path.Combine(dcp1.檔案來源資料夾, "09655588883.si"));
  
            ///////////先處理子文09655588885//////////////
            dcp1 = new MakeDCP("09655588885", sr, null, "3份來文彙辦處理", FlowInfo.FlowType.彙併辦);
            dcp1.檔案來源資料夾 = @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\88885";
            att.Clear(); 
            att.Add("來文附件一.doc");
            dcp1.產生來文文件夾("0970000024", "函", "0970000024.di", att);
            //利用webservice找出彙併辦
            dcp1.設定彙併辦("09655588882", childs);
            dcp1.異動封裝檔("731022", false);
            dcp1.封裝檔XML.Save(Path.Combine(dcp1.檔案來源資料夾, "09655588885.si"));

            ///////////////////處理母文///////////////////
            //傳送者
            //ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "14", "專員", "Jack", "Jack");

            //接收者
            ModifyInfo.signer re = new ModifyInfo.signer("秘書室", "14", "科長", "DSIC", "DSIC");

            //建立封裝類別
            MakeDCP mdcp = new MakeDCP("09655588882", sr, re, "承辦意見", FlowInfo.FlowType.呈核);

            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";

            //所有要加進封裝的檔案存放路徑
            mdcp.檔案來源資料夾 = @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\88882";

            //準備要加入封裝的檔案
            att.Clear(); 
            att.Add("來文附件1.txt");
            att.Add("來文附件2.docx");
            string rtnmsg = mdcp.產生來文文件夾("099100000", "函", "200707020001.di", att);
            att.Clear();
            att.Add("公司OPV1.txt");
            rtnmsg += mdcp.加入文書封裝("便簽", "簽", "100JP00016V1.di", att);
            mdcp.設定彙併辦("09655588882", childs);
            //更新si封裝描述檔,不轉頁面
            if (rtnmsg == "")
            {
                rtnmsg = mdcp.異動封裝檔("731022", false);
            }
            if (rtnmsg != "")
                MessageBox.Show(rtnmsg);
            else
            {
                //將簽核意見htm複製到來資料夾一起壓縮
                File.Copy(@"C:\Program Files\DigitWare\NCC-WEBOE\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾, "100JP00016DcpSignInfo.htm"), true);

                //將所有封裝內容做壓縮
                fileCompress(mdcp);

                //整併後子按檔案上傳後就不需異動除非有解併(回承辦做),往後辦文只需下載母文的壓縮檔即可
            }

        }

        private void 解併_Click(object sender, EventArgs e)
        {
            //解併須回承辦人執行,將母文及要解併的子文下載至client處理

            ///////////////先處理母文//////////////////
            ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "14", "專員", "Jack", "Jack");
            ModifyInfo.signer re = new ModifyInfo.signer("秘書室", "14", "科長", "DSIC", "DSIC");

            FileInfo finfo = new FileInfo(@"D:\temp\09655588882.dsicdp");
            DirectoryInfo dinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\88882");
            //建立封裝類別
            MakeDCP mdcp = new MakeDCP(finfo,dinfo , sr, re, "解併1份子文(0970000024)", FlowInfo.FlowType.解併);

            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            att.Add("來文附件1.txt");
            att.Add("來文附件2.docx");
            string rtnmsg = mdcp.產生來文文件夾("099100000", "函", "200707020001.di", att);
            att.Clear();
            att.Add("公司OPV1.txt");
            rtnmsg += mdcp.加入文書封裝("便簽", "簽", "100JP00016V1.di", att);
            //子文變一份
            string[] childs = new string[1] { "09655588883" };
            mdcp.設定彙併辦("09655588882", childs);
            //更新si封裝描述檔,不轉頁面
            if (rtnmsg == "")
            {
                rtnmsg = mdcp.異動封裝檔("731022", false);
            }
            if (rtnmsg != "")
                MessageBox.Show(rtnmsg);
            else
            {
                //將簽核意見htm複製到來資料夾一起壓縮
                File.Copy(@"C:\Program Files\DigitWare\NCC-WEBOE\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾, "100JP00016DcpSignInfo.htm"), true);
                //將所有封裝內容做壓縮
                fileCompress(mdcp);
            }

            ///////////////////處理要解併的子文////////////////////
            FileInfo cfinfo = new FileInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\88885\09655588885.si");
            DirectoryInfo cdinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\88885");
            MakeDCP dcp1 = new MakeDCP(cfinfo, cdinfo, sr, null, "解併1份子文", FlowInfo.FlowType.解併);
            att.Clear();
            att.Add("來文附件一.doc");
            rtnmsg = dcp1.產生來文文件夾("0970000024", "函", "0970000024.di", att);
            //子文解併不需設定母文資訊
            //dcp1.設定彙併辦("09655588882", childs);
            rtnmsg += dcp1.異動封裝檔("731022", false);
            dcp1.封裝檔XML.Save(Path.Combine(dcp1.檔案來源資料夾, "09655588885.si"));

            if (rtnmsg != "")
                MessageBox.Show(rtnmsg);
            else
                MessageBox.Show("解併成功");

        }

        private void 影音封裝_Click(object sender, EventArgs e)
        {
            MediaPackage mp = new MediaPackage("1000000016", @"C:\WEBOE\DigitalSealed\testDigitalSealed\影音封裝\");
            //string rtnmsg = mp.MakeMediaPackage(xmlTool.MakeNode("詮釋資料", ""), "12345678", @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\影音封裝\output\");
            string rtnmsg = mp.MakeMediaPackage(@"c:\WEBOE\DigitalSealed\testDigitalSealed\影音封裝\output");
            if (rtnmsg != "")
                MessageBox.Show(rtnmsg);
            else
            {
                MessageBox.Show("影音封裝成功");
            }
        }

        private void 借卡_Click(object sender, EventArgs e)
        {
            DCPWS.DCPWSE wse = new DCPWS.DCPWSE();
            XmlNode tmp = new XmlDocument().CreateElement("UserTempCaDetail");//Flow_40_楊啟仁_呈核
            string xml = "<UseOrgSeqNO>1</UseOrgSeqNO><DocNO>1004000014</DocNO><FlowId>Flow_40_楊啟仁_呈核</FlowId><RpsDepID>40</RpsDepID><RpsDepNam>綜合企劃處</RpsDepNam><RpsDivID>4005</RpsDivID><RpsDivNam>電腦資訊科</RpsDivNam><RpsUserID>VICTOR</RpsUserID><RpsUserNam>楊啟仁</RpsUserNam><RpsTime>100/04/28 14:33:45</RpsTime><ApplyDesc>忘記帶</ApplyDesc>";
            tmp.InnerXml = xml;
            wse.updateTempCaDetail(tmp);
 
            //傳送者
            ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "14", "專員", "Jack", "Jack");
            sr.setUseTempCard("10003071151", "申請卡片中"); 

            //接收者
            ModifyInfo.signer re = new ModifyInfo.signer("秘書室", "14", "科長", "DSIC", "DSIC");

            //建立封裝類別
            MakeDCP mdcp = new MakeDCP("1000000016", sr, re, "承辦意見", FlowInfo.FlowType.呈核);

            //轉影像檔時的暫存路徑
            mdcp.暫存資料夾 = @"c:\temp";

            //所有要加進封裝的檔案存放路徑
            mdcp.檔案來源資料夾 = @"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\借卡";

            //準備要加入封裝的檔案
            List<string> att = new List<string>();
            att.Add("來文附件1.txt");
            att.Add("來文附件2.docx");
            string rtnmsg = mdcp.產生來文文件夾("099100000", "函", "200707020001.di", att);
            att.Clear();
            att.Add("公司OPV1.txt");
            rtnmsg += mdcp.加入文書封裝("便簽", "簽", "100JP00016V1.di", att);

            //更新si封裝描述檔,不轉頁面
            if (rtnmsg == "")
            {
                rtnmsg = mdcp.異動封裝檔("731022", false);
            }
            if (rtnmsg != "")
                MessageBox.Show(rtnmsg);
            else
            {
                //將簽核意見htm複製到來資料夾一起壓縮
                File.Copy(@"C:\Program Files\DigitWare\NCC-WEBOE\DcpSignInfo.htm", Path.Combine(mdcp.檔案來源資料夾, "100JP00016DcpSignInfo.htm"), true);
                //將所有封裝內容做壓縮
                fileCompress(mdcp);
            }
        }

        private void 補簽_Click(object sender, EventArgs e)
        {
            //傳送者
            ModifyInfo.signer sr = new ModifyInfo.signer("秘書室", "78", "特助", "ken", "ken");
            //上一流程所產生的封裝壓縮檔位置,來源為fileserver上的文號資料夾
            FileInfo finfo = new FileInfo(@"D:\temp\1000000016.dsicdp");
            //解壓縮上一簽核點所有檔的存放資料夾,通常等於目前簽核點的來源資料夾
            DirectoryInfo dinfo = new DirectoryInfo(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\resign");

            //MakeDCP mdcp = new MakeDCP(finfo, dinfo, sr, null, "", FlowInfo.FlowType.補簽);
            //可由DCPWebService取得要補簽的siXML
            XmlDocument doc = new XmlDocument();
            doc.XmlResolver = null;
            
            //doc.Load(@"c:\WEBOE\DigitalSealed\testDigitalSealed\借卡\1000000016.si");
            doc.Load(@"c:\1000000016.si");
            doc.InsertBefore(doc.CreateXmlDeclaration("1.0", "utf-8", null),doc.FirstChild);
            doc.PreserveWhitespace = true;
            //補簽封裝建構子
            MakeDCP mdcp = new MakeDCP(doc, sr);
            //補簽簽核點由OP傳入
            //string rtn = mdcp.補簽("ds_Flow_01_帝緯系統整合_呈核", "731022", "離職代簽");
            string rtn = mdcp.補簽("Flow_14_Jack_呈核", "731022", "離職代簽");
            if (rtn != "")
                MessageBox.Show(rtn);
            else
                MessageBox.Show("補簽成功");
            mdcp.封裝檔XML.Save(Path.Combine(@"c:\WEBOE\DigitalSealed\testDigitalSealed\resign", "1000000016.si"));
            //將補簽完的si檔利用webservice傳回fileserver的ReSign資料夾
           // DCPWS.DCPWSE wse = new DCPWS.DCPWSE();
            //bool wsertn = wse.SaveReSignXML("1000000016", "sign_Flow_14_Jack_呈核", mdcp.封裝檔XML);
        }

        private void 檔管點收_Click(object sender, EventArgs e)
        {
            //檢測si檔的簽章是否正確

            ////1.將簽核電子檔(si)的<線上簽核>複製到新產生的電子封裝檔////
            Archivist FileOA = new Archivist(@"C:\Users\jack\Desktop\文書規範補正送出公布版作業版\完整版_單層式\R01-200-電子檔案點收資料\1000000025\1000000025.SI", @"C:\Users\jack\Desktop\25.xml");
            string rtn = FileOA.檔管點收加機關憑證("731022");

            //if (rtn != "")
            //    MessageBox.Show(rtn);
            //else
            //    MessageBox.Show("點收成功");

            ////2.使用HSM機關憑證加簽////
            //Archivist FileOA = new Archivist(@"C:\Users\jack\Desktop\OA\1000891006.si");
            //DCPWS.DCPWSE wse = new DCPWS.DCPWSE();
            //string data = FileOA.取得檔管點收加簽前資料();
            //string rtn = wse.SignWithHSM(data, "CheckSignGCA", "mof");
            //if (rtn.IndexOf("err:") > -1)
            //{
            //    MessageBox.Show(rtn);
            //}
            //else
            //{
            //    XmlDocument xdoc = new XmlDocument();
            //    xdoc.XmlResolver = null;
            //    xdoc.LoadXml(rtn);

            //    //File.WriteAllText(@"C:\Users\jack\Desktop\OA\after\1000891006.xml", rtn, Encoding.UTF8);
            //}

            
        }

        private void 詮釋資料_Click(object sender, EventArgs e)
        {
            //Archivist FileOA = new Archivist(@"C:\Users\jack\Desktop\DCP\1000044799_MetaData.xml");
            ////Archivist FileOA = new Archivist(@"D:\WEBOE\拆版專用\DigitalSealed\testDigitalSealed\檔管點收\1000000016.xml","c:\\文號.xml");
            //XmlNode descnd = xmlTool.MakeNode("詮釋資料", "");
            //descnd.InnerXml = "<案件><案由>函調貴部99年4月14日台財訴字第09900074680號訴願卷宗，用畢即行檢還，請查照。</案由><並列案由 /><其他案由>檢送朱豐明君因89年度贈與稅事件本部訴願決定原卷乙宗。</其他案由><主要發文者>財政部</主要發文者><主要來文者>高雄高等行政法院</主要來文者><受文者>高雄高等行政法院</受文者><次要發文者 /><次要來文者 /><文別><公文類別 /></文別><本別 /><密等 /><解密條件或保密期限 /><解密日期 /><保存年限>10</保存年限><發文字號><文字>台財訴字第10000447990號</文字></發文字號><收文字號><文字>1000044799</文字></收文字號><來文字號><文字>高行瓊紀審三100再00058字第1000006140號</文字></來文字號><年度號>0099</年度號><分類號>1551</分類號><案次號>09900250</案次號><卷次號>2</卷次號><目次號>12</目次號><電子或微縮編號 /><其他編號 /><文件產生日期>100/11/07</文件產生日期><檔案類別 /><媒體型式 代碼=\"1\" /><數量 NUM=\"1\">5</數量><計量單位 NUM=\"1\">頁</計量單位><有關案件>測試</有關案件><案名>朱豐明君訴願案</案名><機關代碼>307000000D</機關代碼><單位代碼>I0</單位代碼><電子檔案名稱>1000044799.xml</電子檔案名稱><電子檔案確認日期 /><電子檔案產生者及修改者>張獻聰</電子檔案產生者及修改者><詮釋資料建立者及修改者>王新淳</詮釋資料建立者及修改者></案件>";
            //string rtn = FileOA.加入詮釋資料(descnd, "731022");
            //if (rtn != "")
            //    MessageBox.Show(rtn);
            //else
            //    MessageBox.Show("詮釋資料加入成功");

            ////2.使用HSM機關憑證加簽////
            Archivist FileOA = new Archivist(@"C:\Users\jack\Desktop\OAb\1000000016.xml", "c:\\文號.xml");
            XmlNode descnd = xmlTool.MakeNode("詮釋資料", "");
            descnd.InnerXml = "<案件><案由>供移交匯入用案件三</案由><並列案由 /><其他案由 /><主要發文者>教育部</主要發文者><主要來文者>交通部</主要來文者><受文者 /><文別><公文類別 代碼=\"2\" /><函類別 代碼=\"1\" /></文別><本別 代碼=\"1\"></本別><密等></密等><解密條件或保密期限 /><解密日期 /><保存年限>99</保存年限><發文字號><文字>教字第1000000003號</文字></發文字號><收文字號><文字>1000000003</文字></收文字號><來文字號><文字>交字第1000000003號</文字></來文字號><年度號>100</年度號><分類號>010102</分類號><案次號>001</案次號><卷次號>ED01</卷次號><目次號>001</目次號><文件產生日期>1000121</文件產生日期><檔案類別 代碼=\"2\"></檔案類別><機關代碼>341020000A</機關代碼><電子檔案名稱 /><電子檔案確認日期 /></案件>";
            string rtn = FileOA.加入詮釋資料(descnd, ""); //使用HSM時密碼給空值
            if (rtn != "")
                MessageBox.Show(rtn);
            else
            {
                //經由HSM作加簽
                DCPWS.DCPWSE wse = new DCPWS.DCPWSE();
                XmlDocument xdoc = new XmlDocument();
                xdoc.XmlResolver = null;
                xdoc.Load("c:\\文號.xml");
                string sID = xdoc.DocumentElement.SelectSingleNode("//封裝檔電子簽章/Signature").Attributes["Id"].Value;
                rtn = wse.SignWithHSM(xdoc.OuterXml, sID, "mof");
                if (rtn.IndexOf("err:") > -1)
                {
                    MessageBox.Show(rtn);
                }
                else
                {
                    File.WriteAllText(@"D:\out.xml", rtn, Encoding.UTF8); //寫出最後加簽的XML檔
                }
            }
        }

        private void 數位信封_Click(object sender, EventArgs e)
        {

        }

        private void 移轉封裝檔_Click(object sender, EventArgs e)
        {
            //卡式憑證
            ETransfer et = new ETransfer(@"F:\0000000086\1020802_V001\Media_EVE.xml");
            string[] medianum = new string[1] { "1020802_V001" };
            XmlDocument rtndoc = et.MakeTransferXml(medianum,"731022");
            rtndoc.Save(@"F:\trans.xml");


            //HSM//
            //ETransfer et = new ETransfer(@"C:\Users\jack\Desktop\文書規範補正送出公布版作業版\可正確匯入版本\out.xml");
            ////電子媒體編號
            //string[] medianum = new string[2] { "123456789" , "987654321"};
            //string rtndoc = et.MakeTransferXml(medianum, "").OuterXml;
            //DCPWS.DCPWSE wse = new DCPWS.DCPWSE();
            //string rtn = wse.SignWithHSM(data, et.移轉交封裝檔SigID, "mof");
            //if (rtn.IndexOf("err:") > -1) //發生錯誤
            //{
            //    MessageBox.Show(rtn);
            //}
            //else
            //{
            //    File.WriteAllText(@"D:\out.xml", rtn, Encoding.UTF8); //寫出最後加簽的XML檔
            //}
        }

        private void 媒體封裝檔_Click(object sender, EventArgs e)
        {
            //卡式憑證//
            EMedia mediaxml = new EMedia(@"F:\0000000086\1020802_V001\Media_MetaData.xml", @"F:\TEST.cer");
            mediaxml.MakeEMediaXml("").Save(@"F:\0000000086\1020802_V001\Media_EVE.xml");
            //mediaxml.MakeEMediaXml("1022").Save(@"c:\media.xml");

            //HSM//
            //EMedia mediaxml = new EMedia(@"Q:\加密\Media_MetaData.xml", @"Q:\dsic_1.cer");
            //string data = mediaxml.MakeEMediaXml("").OuterXml;
            //DCPWS.DCPWSE wse = new DCPWS.DCPWSE();
            //string rtn = wse.SignWithHSM(data, mediaxml.媒體封裝SignID , "mof");
            //if (rtn.IndexOf("err:") > -1) //發生錯誤
            //{
            //    MessageBox.Show(rtn);
            //}
            //else
            //{
            //    File.WriteAllText(@"D:\out.xml", rtn, Encoding.UTF8); //寫出最後加簽的XML檔
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog abd = new FolderBrowserDialog();
            if (abd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Signature verifysig = null;
                DirectoryInfo di = new DirectoryInfo(abd.SelectedPath);
                using (StreamWriter sw = new StreamWriter(Path.Combine(abd.SelectedPath,"result.log") , false, Encoding.Default))
                {
                    foreach (FileInfo f in di.GetFiles("*.SI", SearchOption.AllDirectories))
                    {
                        verifysig = new Signature(f.DirectoryName, f.Name);
                        List<string> rtn = verifysig.VerifyXmlFile();
                        foreach (string err in rtn)
                        {
                            sw.Write(err+"\r\n");
                        }
                    }
                }
            }
            MessageBox.Show("完成");
            //Signature verifysig = new Signature(@"C:\Users\jack\Desktop\文書規範補正送出公布版作業版\完整版_單層式\R03-490-電子檔案移交之初始資料\正確\案件56~59\1000000001", "1000000001_c.xml"); 
            //Signature verifysig = new Signature(@"C:\Users\jack\Desktop\OA", "1010690007.si");
            ////List<string> rtn = verifysig.VerifyLastNode(); 
            //List<string> rtn = verifysig.VerifyXmlFile();
            //string msg = "";
            //foreach (string err in rtn)
            //    msg += err;
           // MessageBox.Show(msg == "" ? "檢測成功!" : msg);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EMedia em = new EMedia(@"C:\1011128_V001\Media_EVE.xml", "", "690201");
            //em.DecryptFile(@"C:\Users\jack\Desktop\0000000003\1000523_V001\310390000Q\100=M3B05=001\100=M3B05=001=101=001\1000000022\1000000022.tif", @"C:\Users\jack\Desktop\0000000003\1000523_V001\310390000Q\100=M3B05=001\100=M3B05=001=101=001\1000000022\1000000022_dec.tif"); 
            em.DecryptFile(@"C:\1011128_V001\307000000D\070=N2T01=001\070=N2T01=001=001=001\0700000004\0700000004_D07000000040_0_1.pdf", @"C:\1011128_V001\307000000D\070=N2T01=001\070=N2T01=001=001=001\0700000004\0700000004_D07000000040_0_1_dec.pdf"); 
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //TiffManager tm = new TiffManager(@"C:\Users\jack\Desktop\NTH\1010011575\unzip\101EN00706.tif");
            //System.Collections.ArrayList al = tm.SplitTiffImage("f:\\tmp\\", System.Drawing.Imaging.EncoderValue.CompressionLZW);
        }

        private void mobileBtn_Click(object sender, EventArgs e)
        {
            MobileDCP mDcp = new MobileDCP(@"http://192.168.1.14/mofdcpserver/dcpwse.asmx", @"f:\1010600001.dsicdp", "307000000D", "1010600001", "101E000005");
            string rtn = mDcp.MakeDCP("管理室", "MIB", "警衛", "藍波", "JRNIEH", "萬夫莫敵", "");
            if (!string.IsNullOrEmpty(rtn))
                MessageBox.Show(rtn);
            else
            {
                //mDcp.ForSignXML;待加簽字串
                MessageBox.Show("Finish");
            }
        }

    }
}
