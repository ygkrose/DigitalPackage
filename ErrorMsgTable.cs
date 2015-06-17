using System;
using System.Collections.Generic;
using System.Text;

namespace DigitalSealed
{
    public static class ErrorMsgTable
    {
        /// <summary>
        /// 錯誤代碼清單
        /// </summary>
        /// <param name="ErrCode"></param>
        /// <returns></returns>
        public static string GetErrMessage(int ErrCode)
        {
            string ReturnValue = "";

            switch (ErrCode)
            {
                case 5001:
                    ReturnValue = "一般錯誤";
                    break;
                case 5002:
                    ReturnValue = "配置記憶體錯誤";
                    break;
                case 5005:
                    ReturnValue = "錯誤參數";
                    break;
                case 5007:
                    ReturnValue = "試用版期限已過";
                    break;
                case 5008:
                    ReturnValue = "Base64編碼錯誤";
                    break;
                case 5010:
                    ReturnValue = "無法在MS CryptoAPI Database 中找到指定憑證";
                    break;
                case 5011:
                    ReturnValue = "憑證已過期";
                    break;
                case 5012:
                    ReturnValue = "憑證尚未合法,無法使用";
                    break;
                case 5013:
                    ReturnValue = "憑證可能過期或無法使用";
                    break;
                case 5014:
                    ReturnValue = "憑證主旨錯誤";
                    break;
                case 5015:
                    ReturnValue = "無法找到憑證發行者";
                    break;
                case 5016:
                    ReturnValue = "不合法的憑證簽章";
                    break;
                case 5017:
                    ReturnValue = "憑證用途(加解密,簽驗章)不合適";
                    break;
                case 5020:
                    ReturnValue = "憑證已撤銷";
                    break;
                case 5021:
                    ReturnValue = "憑證已撤銷(金鑰洩露)";
                    break;
                case 5022:
                    ReturnValue = "憑證已撤銷(CA compromised)";
                    break;
                case 5023:
                    ReturnValue = "憑證已撤銷(聯盟已變更)";
                    break;
                case 5024:
                    ReturnValue = "憑證已撤銷(已取代)";
                    break;
                case 5025:
                    ReturnValue = "憑證已撤銷(已停止)";
                    break;
                case 5026:
                    ReturnValue = "憑證保留或暫禁";
                    break;
                case 5028:
                    ReturnValue = "憑證已被CRL撤銷";
                    break;
                case 5030:
                    ReturnValue = "CRL已過期";
                    break;
                case 5031:
                    ReturnValue = "不合法的CRL";
                    break;
                case 5032:
                    ReturnValue = "無法找到CRL";
                    break;
                case 5034:
                    ReturnValue = "CRL簽章值錯誤";
                    break;
                case 5035:
                    ReturnValue = "Digest錯誤";
                    break;
                case 5036:
                    ReturnValue = "不合法的簽章";
                    break;
                case 5037:
                    ReturnValue = "內容錯誤";
                    break;
                case 5040:
                    ReturnValue = "憑證格式錯誤";
                    break;
                case 5041:
                    ReturnValue = "CRL格式錯誤";
                    break;
                case 5042:
                    ReturnValue = "錯誤的PKCS7格式";
                    break;
                case 5043:
                    ReturnValue = "KEY的格式錯誤";
                    break;
                case 5044:
                    ReturnValue = "不合法的PKCS10格式";
                    break;
                case 5045:
                    ReturnValue = "不合適的格式";
                    break;
                case 5046:
                    ReturnValue = "不合法P12格式";
                    break;
                case 5050:
                    ReturnValue = "找不到內部 Reference ID";
                    break;
                case 5060:
                    ReturnValue = "錯誤的憑證或金鑰";
                    break;
                case 5061:
                    ReturnValue = "簽章失敗";
                    break;
                case 5062:
                    ReturnValue = "驗章失敗";
                    break;
                case 5063:
                    ReturnValue = "加密失敗";
                    break;
                case 5064:
                    ReturnValue = "解密失敗";
                    break;
                case 5065:
                    ReturnValue = "產生金鑰失敗";
                    break;
                case 5066:
                    ReturnValue = "刪除憑證失敗";
                    break;
                case 5070:
                    ReturnValue = "使用者按下取消按鈕";
                    break;
                case 5071:
                    ReturnValue = "密碼錯誤";
                    break;
                case 5080:
                    ReturnValue = "無法剖析XML文件";
                    break;
                case 5081:
                    ReturnValue = "無法在XML中,找到指定的標籤名稱";
                    break;
                case 5201:
                    ReturnValue = "開啟STORE失敗";
                    break;
                case 5202:
                    ReturnValue = "憑證鏈不存在";
                    break;
                case 5203:
                    ReturnValue = "無法開啟CSP";
                    break;
                case 5204:
                    ReturnValue = "找不到關連的私密金鑰";
                    break;
                case 5205:
                    ReturnValue = "關連的私密金鑰已被標記成不能匯出";
                    break;
                case 5206:
                    ReturnValue = "Store存取被拒";
                    break;
                case 5901:
                    ReturnValue = "Unicode錯誤";
                    break;
                case 5902:
                    ReturnValue = "找不到外部參考檔案";
                    break;
                case 5904:
                    ReturnValue = "找不到網路路徑";
                    break;
                case 5905:
                    ReturnValue = "錯誤的使用者或密碼";
                    break;
                case 5906:
                    ReturnValue = "執行被拒或沒有權限";
                    break;
                case 9001:
                    ReturnValue = "中斷";
                    break;
                case 9002:
                    ReturnValue = "記憶體錯誤";
                    break;
                case 9003:
                    ReturnValue = "Slot Id不存在";
                    break;
                case 9004:
                    ReturnValue = "一般錯誤";
                    break;
                case 9005:
                    ReturnValue = "函數失敗";
                    break;
                case 9014:
                    ReturnValue = "資料不正確";
                    break;
                case 9016:
                    ReturnValue = "找不到憑證";
                    break;
                case 9018:
                    ReturnValue = "裝置已拔除";
                    break;
                case 9020:
                    ReturnValue = "被加密資料長度錯誤";
                    break;
                case 9024:
                    ReturnValue = "指定的金鑰不存在";
                    break;
                case 9034:
                    ReturnValue = "指定的機制不存在";
                    break;
                case 9036:
                    ReturnValue = "指定的物件不存在";
                    break;
                case 9039:
                    ReturnValue = "PIN碼錯誤";
                    break;
                case 9040:
                    ReturnValue = "Pkcs#11 PIN碼未設定";
                    break;
                case 9041:
                    ReturnValue = "PIN碼錯誤";
                    break;
                case 9043:
                    ReturnValue = "PIN碼錯誤次數超過裝置設定次數,被鎖卡,請先解卡再使用";
                    break;
                case 9044:
                    ReturnValue = "與裝置的連線結束";
                    break;
                case 9045:
                    ReturnValue = "與裝置的連線次數";
                    break;
                case 9046:
                    ReturnValue = "指定的連線不存在";
                    break;
                case 9049:
                    ReturnValue = "指定的連線已存在";
                    break;
                case 9056:
                    ReturnValue = "找不到憑證，請檢查讀卡機跟自然人憑證是否已接上電腦！";
                    break;
                case 9058:
                    ReturnValue = "對裝置作寫的動作時所使用的權限錯誤，可能是SO Pin錯誤或指定的裝置不可寫入";
                    break;
                case 9062:
                    ReturnValue = "已登入裝置";
                    break;
                case 9063:
                    ReturnValue = "未登入裝置";
                    break;
                case 9075:
                    ReturnValue = "緩衝區不夠";
                    break;
                case 9080:
                    ReturnValue = "找不到憑證，請檢查讀卡機跟自然人憑證是否已接上電腦！";
                    break;
                case 9100:
                    ReturnValue = "指定的物件不存在";
                    break;
                case 9101:
                    ReturnValue = "指定的物件已存在";
                    break;
                case 9102:
                    ReturnValue = "裝置中有兩個相同的物件";
                    break;

                default:
                    return ErrCode.ToString();
                    break;
            }
            return ReturnValue;
        }
    }
}
