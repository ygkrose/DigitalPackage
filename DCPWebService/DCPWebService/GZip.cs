using System;
using System.Data;
using System.Configuration;
using System.IO;
using System.Text;
using System.IO.Compression;
using System.Collections;
using System.Collections.Generic;

namespace DCPWebService
{
    public static class GZip
    {

        /// <summary>
        /// Compress
        /// </summary>
        /// <param name="lpSourceFolder">The location of the files to include in the zip file, all files including files in subfolders will be included.</param>
        /// <param name="lpDestFolder">Folder to write the zip file into</param>
        /// <param name="zipFileName">Name of the zip file to write</param>
        public static GZipResult Compress(string lpSourceFolder, string lpDestFolder, string zipFileName)
        {
            return Compress(lpSourceFolder, "*.*", SearchOption.AllDirectories, lpDestFolder, zipFileName, true);
        }

        /// <summary>
        /// Compress
        /// </summary>
        /// <param name="lpSourceFolder">The location of the files to include in the zip file</param>
        /// <param name="searchPattern">Search pattern (ie "*.*" or "*.txt" or "*.gif") to idendify what files in lpSourceFolder to include in the zip file</param>
        /// <param name="searchOption">Only files in lpSourceFolder or include files in subfolders also</param>
        /// <param name="lpDestFolder">Folder to write the zip file into</param>
        /// <param name="zipFileName">Name of the zip file to write</param>
        /// <param name="deleteTempFile">Boolean, true deleted the intermediate temp file, false leaves the temp file in lpDestFolder (for debugging)</param>
        public static GZipResult Compress(string lpSourceFolder, string searchPattern, SearchOption searchOption, string lpDestFolder, string zipFileName, bool deleteTempFile)
        {
            DirectoryInfo di = new DirectoryInfo(lpSourceFolder);
            FileInfo[] files = di.GetFiles("*.*", searchOption);
            return Compress(files, lpSourceFolder, lpDestFolder, zipFileName, deleteTempFile);
        }

        /// <summary>
        /// Compress
        /// </summary>
        /// <param name="files">Array of FileInfo objects to be included in the zip file</param>
        /// <param name="lpBaseFolder">Base folder to use when creating relative paths for the files 
        /// stored in the zip file. For example, if lpBaseFolder is 'C:\zipTest\Files\', and there is a file 
        /// 'C:\zipTest\Files\folder1\sample.txt' in the 'files' array, the relative path for sample.txt 
        /// will be 'folder1/sample.txt'</param>
        /// <param name="lpDestFolder">Folder to write the zip file into</param>
        /// <param name="zipFileName">Name of the zip file to write</param>
        public static GZipResult Compress(FileInfo[] files, string lpBaseFolder, string lpDestFolder, string zipFileName)
        {
            return Compress(files, lpBaseFolder, lpDestFolder, zipFileName, true);
        }

        public static GZipResult Compress(List<FileInfo> files, string lpBaseFolder, string lpDestFolder, string zipFileName)
        {
            FileInfo[] filesinfo = new FileInfo[files.Count]; 
            for (int i = 0; i < files.Count; i++)
            {
                filesinfo[i] = files[i]; 
            }
            return Compress(filesinfo, lpBaseFolder, lpDestFolder, zipFileName, true);
        }

        /// <summary>
        /// Compress
        /// </summary>
        /// <param name="files">Array of FileInfo objects to be included in the zip file</param>
        /// <param name="lpBaseFolder">Base folder to use when creating relative paths for the files 
        /// stored in the zip file. For example, if lpBaseFolder is 'C:\zipTest\Files\', and there is a file 
        /// 'C:\zipTest\Files\folder1\sample.txt' in the 'files' array, the relative path for sample.txt 
        /// will be 'folder1/sample.txt'</param>
        /// <param name="lpDestFolder">Folder to write the zip file into</param>
        /// <param name="zipFileName">Name of the zip file to write</param>
        /// <param name="deleteTempFile">Boolean, true deleted the intermediate temp file, false leaves the temp file in lpDestFolder (for debugging)</param>
        public static GZipResult Compress(FileInfo[] files, string lpBaseFolder, string lpDestFolder, string zipFileName, bool deleteTempFile)
        {
            GZipResult result = new GZipResult();

            if (!lpDestFolder.EndsWith("\\"))
            {
                lpDestFolder += "\\";
            }

            string lpTempFile = lpDestFolder + zipFileName + ".tmp";
            string lpZipFile = lpDestFolder + zipFileName;

            result.TempFile = lpTempFile;
            result.ZipFile = lpZipFile;

            int fileCount = 0;

            if (files != null && files.Length > 0)
            {
                CreateTempFile(files, lpBaseFolder, lpTempFile, result);

                if (result.FileCount > 0)
                {
                    CreateZipFile(lpTempFile, lpZipFile, result);
                }

                // delete the temp file
                try
                {
                    if (deleteTempFile)
                    {
                        File.Delete(lpTempFile);
                        result.TempFileDeleted = true;
                    }
                }
                catch (Exception ex4)
                {
                    // handle or display the error 
                    throw ex4;
                }
            }
            return result;
        }

        private static void CreateZipFile(string lpSourceFile, string lpZipFile, GZipResult result)
        {
            byte[] buffer;
            int count = 0;
            FileStream fsOut = null;
            FileStream fsIn = null;
            GZipStream gzip = null;

            // compress the file into the zip file
            try
            {
                fsOut = new FileStream(lpZipFile, FileMode.Create, FileAccess.Write, FileShare.None);
                gzip = new GZipStream(fsOut, CompressionMode.Compress, true);

                fsIn = new FileStream(lpSourceFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                buffer = new byte[fsIn.Length];
                count = fsIn.Read(buffer, 0, buffer.Length);
                fsIn.Close();
                fsIn = null;

                // compress to the zip file
                gzip.Write(buffer, 0, buffer.Length);

                result.ZipFileSize = fsOut.Length;
                result.CompressionPercent = GetCompressionPercent(result.TempFileSize, result.ZipFileSize);
            }
            catch (Exception ex1)
            {
                // handle or display the error 
                throw ex1;
            }
            finally
            {
                if (gzip != null)
                {
                    gzip.Close();
                    gzip = null;
                }
                if (fsOut != null)
                {
                    fsOut.Close();
                    fsOut = null;
                }
                if (fsIn != null)
                {
                    fsIn.Close();
                    fsIn = null;
                }
            }
        }

        private static void CreateTempFile(FileInfo[] files, string lpBaseFolder, string lpTempFile, GZipResult result)
        {
            byte[] buffer;
            int count = 0;
            byte[] header;
            string fileHeader = null;
            string fileModDate = null;
            string lpFolder = null;
            int fileIndex = 0;
            string lpSourceFile = null;
            string vpSourceFile = null;
            GZippedFile gzf = null;
            FileStream fsOut = null;
            FileStream fsIn = null;

            if (files != null && files.Length > 0)
            {
                try
                {
                    result.Files = new GZippedFile[files.Length];

                    // open the temp file for writing
                    fsOut = new FileStream(lpTempFile, FileMode.Create, FileAccess.Write, FileShare.None);

                    foreach (FileInfo fi in files)
                    {
                        lpFolder = fi.DirectoryName + "\\";
                        try
                        {
                            gzf = new GZippedFile();
                            gzf.Index = fileIndex;

                            // read the source file, get its virtual path within the source folder
                            lpSourceFile = fi.FullName;
                            gzf.LocalPath = lpSourceFile;
                            vpSourceFile = lpSourceFile.Replace(lpBaseFolder, string.Empty);
                            vpSourceFile = vpSourceFile.Replace("\\", "/");
                            gzf.RelativePath = vpSourceFile;

                            fsIn = new FileStream(lpSourceFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                            buffer = new byte[fsIn.Length];
                            count = fsIn.Read(buffer, 0, buffer.Length);
                            fsIn.Close();
                            fsIn = null;

                            fileModDate = fi.LastWriteTimeUtc.ToString();
                            gzf.ModifiedDate = fi.LastWriteTimeUtc;
                            gzf.Length = buffer.Length;

                            fileHeader = fileIndex.ToString() + "," + vpSourceFile + "," + fileModDate + "," + buffer.Length.ToString() + "\n";
                            header = Encoding.UTF8.GetBytes(fileHeader);

                            if (!fileHeader.Equals(System.Text.Encoding.UTF8.GetString(header)))
                            {
                                header = Encoding.Default.GetBytes(fileHeader);
                                if (!fileHeader.Equals(System.Text.Encoding.Default.GetString(header)))
                                {
                                    header = Encoding.GetEncoding("big5").GetBytes(fileHeader);
                                }
                            }


                            fsOut.Write(header, 0, header.Length);
                            fsOut.Write(buffer, 0, buffer.Length);
                            fsOut.WriteByte(10); // linefeed

                            gzf.AddedToTempFile = true;

                            // update the result object
                            result.Files[fileIndex] = gzf;

                            // increment the fileIndex
                            fileIndex++;
                        }
                        catch (Exception ex1)
                        {
                            // handle or display the error 
                            throw ex1;
                        }
                        finally
                        {
                            if (fsIn != null)
                            {
                                fsIn.Close();
                                fsIn = null;
                            }
                        }
                        if (fsOut != null)
                        {
                            result.TempFileSize = fsOut.Length;
                        }
                    }
                }
                catch (Exception ex2)
                {
                    // handle or display the error 
                    throw ex2;
                }
                finally
                {
                    if (fsOut != null)
                    {
                        fsOut.Close();
                        fsOut = null;
                    }
                }
            }

            result.FileCount = fileIndex;
        }

        public static GZipResult Decompress(string lpSourceFolder, string lpDestFolder, string zipFileName)
        {
            return Decompress(lpSourceFolder, lpDestFolder, zipFileName, true);
        }

        public static GZipResult Decompress(string lpSrcFolder, string lpDestFolder, string zipFileName, bool deleteTempFile)
        {
            GZipResult result = new GZipResult();

            if (!lpDestFolder.EndsWith("\\"))
            {
                lpDestFolder += "\\";
            }
            if (!Directory.Exists(lpDestFolder))
                Directory.CreateDirectory(lpDestFolder);  
            string lpTempFile =Path.Combine(lpDestFolder , zipFileName + ".tmp");
            string lpZipFile = Path.Combine(lpSrcFolder ,zipFileName);

            result.TempFile = lpTempFile;
            result.ZipFile = lpZipFile;

            string line = null;
            string lpFilePath = null;
            GZippedFile gzf = null;
            FileStream fsTemp = null;
            ArrayList gzfs = new ArrayList();

            // extract the files from the temp file
            try
            {
                fsTemp = UnzipToTempFile(lpZipFile, lpTempFile, result);
                if (fsTemp != null)
                {
                    while (fsTemp.Position != fsTemp.Length)
                    {
                        line = null;
                        while (string.IsNullOrEmpty(line) && fsTemp.Position != fsTemp.Length)
                        {
                            line = ReadLine(fsTemp);
                        }

                        if (!string.IsNullOrEmpty(line))
                        {
                            gzf = GZippedFile.GetGZippedFile(line);
                            if (gzf != null && gzf.Length > 0)
                            {
                                gzfs.Add(gzf);
                                lpFilePath = lpDestFolder + gzf.RelativePath;
                                gzf.LocalPath = lpFilePath;
                                WriteFile(fsTemp, gzf.Length, lpFilePath);
                                gzf.Restored = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex3)
            {
                // handle or display the error 
                throw ex3;
            }
            finally
            {
                if (fsTemp != null)
                {
                    fsTemp.Close();
                    fsTemp = null;
                }
            }

            // delete the temp file
            try
            {
                if (deleteTempFile)
                {
                    File.Delete(lpTempFile);
                    result.TempFileDeleted = true;
                }
            }
            catch (Exception ex4)
            {
                // handle or display the error 
                throw ex4;
            }

            result.FileCount = gzfs.Count;
            result.Files = new GZippedFile[gzfs.Count];
            gzfs.CopyTo(result.Files);
            return result;
        }

        private static string ReadLine(FileStream fs)
        {
            string line = string.Empty;

            const int bufferSize = 4096;
            byte[] buffer = new byte[bufferSize];
            byte b = 0;
            byte lf = 10;
            int i = 0;

            while (b != lf)
            {
                b = (byte)fs.ReadByte();
                buffer[i] = b;
                i++;
            }

            line = System.Text.Encoding.Default.GetString(buffer, 0, i - 1);
            if (line.IndexOf("?")>-1)
                line = System.Text.Encoding.UTF8.GetString(buffer, 0, i - 1);
            if (line.IndexOf("?") > -1)
                line = System.Text.Encoding.GetEncoding("big5").GetString(buffer, 0, i - 1);
            return line;
        }

        private static void WriteFile(FileStream fs, int fileLength, string lpFile)
        {
            FileStream fsFile = null;

            try
            {
                string lpFolder = GetFolder(lpFile);
                if (!string.IsNullOrEmpty(lpFolder) && lpFolder != lpFile && !Directory.Exists(lpFolder))
                {
                    Directory.CreateDirectory(lpFolder);
                }

                byte[] buffer = new byte[fileLength];
                int count = fs.Read(buffer, 0, fileLength);
                fsFile = new FileStream(lpFile, FileMode.Create, FileAccess.Write, FileShare.None);
                //fsFile.Write(buffer, 0, buffer.Length);
                fsFile.Write(buffer, 0, count);
            }
            catch (Exception ex2)
            {
                // handle or display the error 
                throw ex2;
            }
            finally
            {
                if (fsFile != null)
                {
                    fsFile.Flush();
                    fsFile.Close();
                    fsFile = null;
                }
            }
        }

        private static string GetFolder(string filename)
        {
            string vpFolder = filename;
            int index = filename.LastIndexOf("/");
            if (index != -1)
            {
                vpFolder = filename.Substring(0, index + 1);
            }
            return vpFolder;
        }

        private static FileStream UnzipToTempFile(string lpZipFile, string lpTempFile, GZipResult result)
        {
            FileStream fsIn = null;
            GZipStream gzip = null;
            FileStream fsOut = null;
            FileStream fsTemp = null;

            const int bufferSize = 4096;
            byte[] buffer = new byte[bufferSize];
            int count = 0;

            try
            {
                fsIn = new FileStream(lpZipFile, FileMode.Open, FileAccess.Read, FileShare.Read);
                result.ZipFileSize = fsIn.Length;

                fsOut = new FileStream(lpTempFile, FileMode.Create, FileAccess.Write, FileShare.None);
                gzip = new GZipStream(fsIn, CompressionMode.Decompress, true);
                while (true)
                {
                    count = gzip.Read(buffer, 0, bufferSize);
                    if (count != 0)
                    {
                        fsOut.Write(buffer, 0, count);
                    }
                    if (count != bufferSize)
                    {
                        break;
                    }
                }
            }
            catch (Exception ex1)
            {
                // handle or display the error 
                throw ex1;
            }
            finally
            {
                if (gzip != null)
                {
                    gzip.Close();
                    gzip = null;
                }
                if (fsOut != null)
                {
                    fsOut.Close();
                    fsOut = null;
                }
                if (fsIn != null)
                {
                    fsIn.Close();
                    fsIn = null;
                }
            }

            fsTemp = new FileStream(lpTempFile, FileMode.Open, FileAccess.Read, FileShare.None);
            if (fsTemp != null)
            {
                result.TempFileSize = fsTemp.Length;
            }
            return fsTemp;
        }

        private static int GetCompressionPercent(long tempLen, long zipLen)
        {
            double tmp = (double)tempLen;
            double zip = (double)zipLen;
            double hundred = 100;

            double ratio = (tmp - zip) / tmp;
            double pcnt = ratio * hundred;

            return (int)pcnt;
        }
    }

    public class GZippedFile
    {
        public int Index = 0;
        public string RelativePath = null;
        public DateTime ModifiedDate;
        public int Length = 0;
        public bool AddedToTempFile = false;
        public bool Restored = false;
        public string LocalPath = null;
        public string Folder = null;

        public static GZippedFile GetGZippedFile(string fileInfo)
        {
            GZippedFile gzf = null;

            if (!string.IsNullOrEmpty(fileInfo))
            {
                // get the file information
                string[] info = fileInfo.Split(',');
                if (info != null && info.Length == 4)
                {
                    gzf = new GZippedFile();
                    gzf.Index = Convert.ToInt32(info[0]);
                    gzf.RelativePath = info[1].Replace("/", "\\");
                    gzf.ModifiedDate = DateTime.Now;// Convert.ToDateTime(info[2]);
                    gzf.Length = Convert.ToInt32(info[3]);
                }
            }
            return gzf;
        }
    }

    public class GZipResult
    {
        public GZippedFile[] Files = null;
        public int FileCount = 0;
        public long TempFileSize = 0;
        public long ZipFileSize = 0;
        public int CompressionPercent = 0;
        public string TempFile = null;
        public string ZipFile = null;
        public bool TempFileDeleted = false;
    }
}

