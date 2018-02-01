using System;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace S1ExcelPlugIn
{
    class Crypto
    {
        public string Encrypt(string decrypted)
        {
            byte[] data = System.Text.ASCIIEncoding.ASCII.GetBytes(decrypted);
            byte[] rgbKey = System.Text.ASCIIEncoding.ASCII.GetBytes("72612512");
            byte[] rgbIV = System.Text.ASCIIEncoding.ASCII.GetBytes("25627182");

            MemoryStream memoryStream = new MemoryStream(1024);

            DESCryptoServiceProvider desCryptoServiceProvider = new
            DESCryptoServiceProvider();

            CryptoStream cryptoStream = new CryptoStream(memoryStream,
            desCryptoServiceProvider.CreateEncryptor(rgbKey, rgbIV),
            CryptoStreamMode.Write);

            cryptoStream.Write(data, 0, data.Length);

            cryptoStream.FlushFinalBlock();

            byte[] result = new byte[(int)memoryStream.Position];

            memoryStream.Position = 0;

            memoryStream.Read(result, 0, result.Length);

            cryptoStream.Close();

            return System.Convert.ToBase64String(result);
        }
        public string Decrypt(string encrypted)
        {
            if (encrypted == null)
            {
                return "";
            }
            else
            {
                byte[] data = System.Convert.FromBase64String(encrypted);
                byte[] rgbKey = System.Text.ASCIIEncoding.ASCII.GetBytes("72612512");
                byte[] rgbIV = System.Text.ASCIIEncoding.ASCII.GetBytes("25627182");

                MemoryStream memoryStream = new MemoryStream(data.Length);

                DESCryptoServiceProvider desCryptoServiceProvider = new
                DESCryptoServiceProvider();

                CryptoStream cryptoStream = new CryptoStream(memoryStream,
                desCryptoServiceProvider.CreateDecryptor(rgbKey, rgbIV),
                CryptoStreamMode.Read);

                memoryStream.Write(data, 0, data.Length);

                memoryStream.Position = 0;

                string decrypted = new StreamReader(cryptoStream).ReadToEnd();

                cryptoStream.Close();
                return decrypted;
            }
        }

        #region Saving and Retrieving from Registry
        public String GetSettings(String subkey, String key)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\" + subkey);
            String keyValue = (String)regKey.GetValue(key);
            regKey.Close();
            return keyValue;
        }

        public int GetSubKeyCount(String subkey)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\" + subkey);
            int count = regKey.ValueCount;
            regKey.Close();
            return count;
        }

        public void SetSettings(String subkey, String key, String value)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\" + subkey);
            regKey.SetValue(key, value);
            regKey.Close();
        }

        #endregion
    }
}
