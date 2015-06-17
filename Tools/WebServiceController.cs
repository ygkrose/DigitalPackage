using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Web.Services.Description;
using System.Web.Services.Protocols;
using System.IO;
using System.CodeDom;
using Microsoft.CSharp;
using System.CodeDom.Compiler;



namespace DSIC.WEBOE.MAIN.Controller
{
    static class WebServiceController
    {
        public static object InvokeWebservice(string pUrl, string @pNamespace, string pClassname, string pMethodname, object[] pArgs)
        {
            WebClient tWebClient = new WebClient();
            //tWebClient.Credentials = System.Net.CredentialCache.DefaultCredentials;
            ServicePointManager.ServerCertificateValidationCallback =  delegate { return true; }; //啟用https時需有此行避免掉檢查憑證的對話框問題
            //讀取WSDL檔，確認Web Service描述內容 
            Stream tStream = tWebClient.OpenRead(pUrl + "?WSDL");
            ServiceDescription tServiceDesp = ServiceDescription.Read(tStream);
            //將讀取到的WSDL檔描述import近來 
            ServiceDescriptionImporter tServiceDespImport = new ServiceDescriptionImporter();
            tServiceDespImport.AddServiceDescription(tServiceDesp, "", "");
            CodeNamespace tCodeNamespace = new CodeNamespace(@pNamespace);
            //指定要編譯程式 
            CodeCompileUnit tCodeComUnit = new CodeCompileUnit();
            tCodeComUnit.Namespaces.Add(tCodeNamespace);
            tServiceDespImport.Import(tCodeNamespace, tCodeComUnit);

            //以C#的Compiler來進行編譯 
            CSharpCodeProvider tCSProvider = new CSharpCodeProvider();
            ICodeCompiler tCodeCom = tCSProvider.CreateCompiler();

            //設定編譯參數 
            System.CodeDom.Compiler.CompilerParameters tComPara = new System.CodeDom.Compiler.CompilerParameters();
            tComPara.GenerateExecutable = false;
            tComPara.GenerateInMemory = true;

            //取得編譯結果 
            System.CodeDom.Compiler.CompilerResults tComResult = tCodeCom.CompileAssemblyFromDom(tComPara, tCodeComUnit);

            //如果編譯有錯誤的話，將錯誤訊息丟出 
            if (true == tComResult.Errors.HasErrors)
            {
                System.Text.StringBuilder tStr = new System.Text.StringBuilder();
                foreach (System.CodeDom.Compiler.CompilerError tComError in tComResult.Errors)
                {
                    tStr.Append(tComError.ToString());
                    tStr.Append(System.Environment.NewLine);
                }
                throw new Exception(tStr.ToString());
            }

            //取得編譯後產出的Assembly 
            System.Reflection.Assembly tAssembly = tComResult.CompiledAssembly;
            Type tType = tAssembly.GetType(@pNamespace + "." + pClassname, true, true);
            object tTypeInstance = Activator.CreateInstance(tType);
            //若WS有overload的話，需明確指定參數內容 
            Type[] tArgsType = null;
            if (pArgs == null)
            {
                tArgsType = new Type[0];
            }
            else
            {
                int tArgsLength = pArgs.Length;
                tArgsType = new Type[tArgsLength];
                for (int i = 0; i < tArgsLength; i++)
                {
                    tArgsType[i] = pArgs[i].GetType();
                }
            }

            //若沒有overload的話，第二個參數便不需要，這邊要注意的是WsiProfiles.BasicProfile1_1本身不支援Web Service overload，因此需要改成不遵守WsiProfiles.BasicProfile1_1協議 
            //System.Reflection.MethodInfo tInvokeMethod = tType.GetMethod(pMethodname, tArgsType);
            System.Reflection.MethodInfo tInvokeMethod = tType.GetMethod(pMethodname);
            //要加這三行：如果是Windows整合驗證的話，透過SoapHttp來對要invoke的目標WS做驗證
            //SoapHttpClientProtocol webRequest = (SoapHttpClientProtocol)tTypeInstance;
            //webRequest.PreAuthenticate = true;
            //webRequest.Credentials = System.Net.CredentialCache.DefaultCredentials;

            //實際invoke該method 
            return tInvokeMethod.Invoke(tTypeInstance, pArgs);
        }
    }
}
