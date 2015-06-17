using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Web.Services;
using System.Web.Services.Protocols;

    /// <summary>
    /// WSEFileDownloadClient 的摘要描述。
    /// </summary>
namespace DCPWebService
{
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name = "PreferencesServiceSoap", Namespace = "http://tempuri.org")]
    public class WSEFileDownloadClient : System.Web.Services.Protocols.SoapHttpClientProtocol
    {
        /// <summary>
        /// 建構.
        /// </summary>
        /// 

        public WSEFileDownloadClient(string url)
        {
            this.Url = url.Trim();

        }

        /// <summary>
        /// 下載檔案.
        /// </summary>
        ///
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/DownloadFile",
             RequestNamespace = "http://tempuri.org/",
             ResponseNamespace = "http://tempuri.org/",
             Use = System.Web.Services.Description.SoapBindingUse.Literal,
             ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public byte[] DownloadFile(string FName)
        {
            object[] results = this.Invoke("DownloadFile", new object[] { FName });

            return ((byte[])results[0]);
        }

        /// <summary>
        /// 開始下載檔案.
        /// </summary>
        /// 

        public System.IAsyncResult BeginDownloadFile(string FName, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("DownloadFile", new object[] { FName }, callback, asyncState);
        }

        /// <summary>
        /// 結束下載檔案.
        /// </summary>
        /// 

        public byte[] EndDownloadFile(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((byte[])results[0]);
        }


        /// <summary>
        /// 檔案上傳.
        /// </summary>
        ///
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/UploadFile",
             RequestNamespace = "http://tempuri.org/",
             ResponseNamespace = "http://tempuri.org/",
             Use = System.Web.Services.Description.SoapBindingUse.Literal,
             ParameterStyle = System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public byte[] UploadFile(string FName)
        {
            object[] results = this.Invoke("UploadFile", new object[] { FName });

            return ((byte[])results[0]);
        }

        /// <summary>
        /// 開始檔案上傳.
        /// </summary>
        /// 

        public System.IAsyncResult BeginUploadFile(string FName, System.AsyncCallback callback, object asyncState)
        {
            return this.BeginInvoke("UploadFile", new object[] { FName }, callback, asyncState);
        }

        /// <summary>
        /// 結束檔案上傳.
        /// </summary>
        /// 

        public byte[] EndUploadFile(System.IAsyncResult asyncResult)
        {
            object[] results = this.EndInvoke(asyncResult);
            return ((byte[])results[0]);
        }
    }

}