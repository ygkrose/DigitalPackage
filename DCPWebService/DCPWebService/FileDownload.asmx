<%@ WebService Language="C#" Class="FileDownload" %>

using System;
using System.IO;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class FileDownload  : System.Web.Services.WebService {

    /// <summary>
    /// 檔下下載
    /// </summary>
    /// <param name="FName"></param>
    /// <returns></returns>
    [WebMethod]
    public byte[] DownloadFile(string FName)
    {
        System.IO.FileStream fs1 = null;
        fs1 = System.IO.File.Open(FName, FileMode.Open, FileAccess.Read);
        byte[] b1 = new byte[fs1.Length];
        fs1.Read(b1, 0, (int)fs1.Length);
        fs1.Close();
        return b1;
    }

    /// <summary>
    /// 檔案上傳
    /// </summary>
    /// <param name="FileContent"></param>
    /// <param name="FName"></param>
    /// <returns></returns>
    [WebMethod]
    public bool UploadFile(byte[] FileContent,string FName)
    {
        try
        {
            //目的目錄
            string Target_Folder = Path.GetDirectoryName(FName);
            if (!Directory.Exists(Target_Folder))//路徑不存在
                Directory.CreateDirectory(Target_Folder);

            //FileContent = DownloadFile(@"D:\Project\NCC\WebService\ftproot\Web\AtaShare\098\1L\00896\0981L008930.tif");
            
            System.IO.FileStream fs1 = new FileStream(FName, FileMode.Create);
            fs1.Write(FileContent, 0, (int)FileContent.Length);
            fs1.Flush();
            fs1.Close();
            return true;
        }
        catch (Exception err)
        {
            return false;
        }
    }
}

