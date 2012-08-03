using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QQGroupMerger {

    /// <summary>
    /// 扩展方法：
    /// 1. 类必须为static
    /// 2. 每个方法必须为public static
    /// 3. 第一个参数必须有this
    /// </summary>
    static class Extendsions {

        /// <summary>
        /// 求字符串中某关键字（第一次出现）前面的部分。如果没有匹配内容则返回""
        /// </summary>
        public static string SubstringBefore(this string str, string key) {
            var index = str.IndexOf(key);
            if (index == -1) return "";
            return str.Substring(0, index);
        }

        /// <summary>
        /// 求字符串中指定的开始（第一次出现）与结束关键字（最后一次出现）中间的部分。如果没有匹配内容则返回""
        /// </summary>
        public static string SubstringBetween(this string str, string start, string end) {
            var index1 = str.IndexOf(start);
            var index2 = str.LastIndexOf(end);
            if (index1 == -1 || index2 == -1) return "";
            return str.Substring(index1 + start.Length, index2 - index1 - start.Length);
        }

        /// <summary>
        /// 求字符串中指定的关键字（第一次出现）之后的部分。如果没有匹配内容则返回""
        /// </summary>
        public static string SubstringAfter(this string str, string key) {
            var index = str.IndexOf(key);
            if (index == -1) return "";
            return str.Substring(index + key.Length);
        }

        /// <summary>
        /// 求出一个字符串的md5摘要
        /// </summary>
        public static string Md5(this string text) {
            // step 1, calculate MD5 hash from input
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(text);
            byte[] hash = md5.ComputeHash(inputBytes);

            // step 2, convert byte array to hex string
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++) {
                sb.Append(hash[i].ToString("X2")); // 16进制2字节
            }
            return sb.ToString();
        }
    }
}
