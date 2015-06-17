<%@ WebService Language="C#" Class="WSEFileMove" %>

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.IO;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class WSEFileMove : System.Web.Services.WebService
{
    [WebMethod]
    public bool WebFileExist(string sDirPath, string sFileName)
    {
        try
        {
            string Dir1 = sDirPath.Substring(sDirPath.Length - 1, 1);
            if (Dir1 != "\\")
            {
                sDirPath += "\\";
            }
            bool blnFileExist;
            blnFileExist = File.Exists(sDirPath + sFileName);
            if (!blnFileExist)
            {
                //throw new Exception("檔案不存在.");
                return false;
            }
        }
        catch
        {
            //throw new Exception("錯誤的檔案格式.");
            return false;
        }
        return true;
    }

    [WebMethod]
    public bool WebCopyFile(string sDir1, string sFileName1, string sDir2, string sFileName2, bool overWrite)
    {
        try
        {
            FileInfo myFileInfo = new FileInfo(sDir1 + sFileName1);
            if (overWrite)
            {
                myFileInfo.CopyTo(sDir2 + sFileName2, true);
            }
            else
            {
                myFileInfo.CopyTo(sDir2 + sFileName2, false);
            }
            return true;
        }
        catch
        {
            return false;
        }
			
    }

    [WebMethod]
    public bool WebCopyPathFile(string sDirFileName1, string sDirFileName2, bool overWrite)
    {
        try
        {
            FileInfo myFileInfo = new FileInfo(sDirFileName1);
            if (overWrite)
            {
                myFileInfo.CopyTo(sDirFileName2, true);
            }
            else
            {
                myFileInfo.CopyTo(sDirFileName2, false);
            }
            return true;
        }
        catch
        {
            return false;
        }

    }

    [WebMethod]
    public bool WebDeleteFile(string sDirPath, string sFileName)
    {
        try
        {
            if (FileExist(sDirPath, sFileName))
            {
                FileInfo myFileInfo = new FileInfo(sDirPath + sFileName);
                myFileInfo.Delete();
            }
            else
            {
                //throw new Exception("檔案不存在.");
                return false;
            }
        }
        catch
        {
            //throw new Exception("刪除檔案錯誤.");
            return false;
        }

        return true;
    }

    [WebMethod]
    public bool WebCreateFileDir(string sDirPath)
    {
        try
        {
            if (!Directory.Exists(sDirPath))
            {

                DirectoryInfo myDirectoryInfo = new DirectoryInfo(sDirPath);
                myDirectoryInfo.Create();
            }
        }
        catch
        {
            //throw new Exception("檔案目錄建立失敗.");
            return false;
        }
        return true;
    }

    [WebMethod]
    public bool WebCreateMuchFileDir(string sDirPath)
    {
        string mDir = "";
        string pathDir = "";
        mDir = sDirPath.Trim();
        int k = 0;

        k = mDir.IndexOf("\\", 0, mDir.Length);
        if (k < 0)
        {
            return false;
        }
        pathDir = mDir.Substring(0, k + 1);
        mDir = mDir.Substring(k + 1, mDir.Length - k - 1);
        while (mDir != "")
        {
            k = mDir.IndexOf("\\", 0, mDir.Length);
            if (k > 0)
            {
                pathDir = pathDir + mDir.Substring(0, k + 1);
                mDir = mDir.Substring(k + 1, mDir.Length - k - 1);
                if (CreateFileDir(pathDir) == false)
                {
                    return false;
                }
            }
            else
            {
                mDir = "";
            }
        }
        return true;
    }
    
    [WebMethod]
    public bool WebDirectoryExists(string sDirPath)
    {
        if (Directory.Exists(sDirPath))
            return true;
        else
            return false;
    }

    [WebMethod]
    public bool WebDirectoryDelete(string sDirPath)
    {
        try
        {
            Directory.Delete(sDirPath);
        }
        catch
        {
            return false;
        }
        return true;
    }

    [WebMethod]
    public bool WebDirectoryMove(string sDirPath, string tDirPath)
    {
        try
        {
            Directory.Move(sDirPath, tDirPath);
            return true;
        }
        catch
        {
            return false;
        }
    }


    [WebMethod]
    public string[] WebDirectoryGetFiles(string sDirPath, string sPattern)
    {
        string[] sFilesNam = { };
        
        try
        {
            if (sPattern.Trim() == "")
            sFilesNam = Directory.GetFiles(sDirPath);          
            else
                sFilesNam = Directory.GetFiles(sDirPath, sPattern);          
              
            
        }
        catch
        {
            return sFilesNam;
        }
        return sFilesNam;
    }
    
    
    public bool FileExist(string strDir, string strFilename)
    {
        try
        {
            string Dir1 = strDir.Substring(strDir.Length - 1, 1);
            if (Dir1 != "\\")
            {
                strDir += "\\";
            }
            bool blnFileExist;
            blnFileExist = File.Exists(strDir + strFilename);
            if (!blnFileExist)
            {
                //throw new Exception("檔案不存在.");
                return false;
            }
        }
        catch
        {
            //throw new Exception("錯誤的檔案格式.");
            return false;
        }
        return true;
    }

    /// <summary>
	/// CreateFileDir 建立目錄。
	/// </summary>
	public bool CreateFileDir(string strDir)
	{
		try
		{
			if (!Directory.Exists(strDir))
			{
				
				DirectoryInfo myDirectoryInfo = new DirectoryInfo(strDir);
				myDirectoryInfo.Create();
			}
		}
		catch
		{
			//throw new Exception("檔案目錄建立失敗.");
			return false;
		}
        return true;
    }
    
}