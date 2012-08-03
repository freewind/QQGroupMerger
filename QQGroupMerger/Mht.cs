using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Globalization;

namespace QQGroupMerger {

    class MhtWriter {
        private string file;
        private string boundary;
        private StreamWriter writer;
        public MhtWriter(string file, string boundary) {
            this.file = file;
            this.boundary = boundary;
        }
        public void writeStart() {
            this.writer = new StreamWriter(file);
            writer.WriteLine(
@"From: <Save by Tencent MsgMgr>
Subject: Tencent IM Message
MIME-Version: 1.0
Content-Type:multipart/related;
	charset=""utf-8""
	type=""text/html"";
	boundary=""" + boundary + "\"");
        }

        public void writeHtml(string html) {
            writer.WriteLine();
            writer.WriteLine("--" + boundary);
            writer.WriteLine("Content-Type: text/html");
            writer.WriteLine("Content-Transfer-Encoding:7bit");
            writer.WriteLine();
            writer.WriteLine(html);
            writer.WriteLine();
        }

        public void writeImage(string imageFile) {
            var filename = Path.GetFileName(imageFile);
            writer.WriteLine("--" + boundary);
            writer.WriteLine("Content-Type:image/" + filename.Split('.')[1]);
            writer.WriteLine("Content-Transfer-Encoding:base64");
            writer.WriteLine("Content-Location:images/" + filename);
            writer.WriteLine();
            var bin = File.ReadAllBytes(imageFile);
            var base64 = System.Convert.ToBase64String(bin);
            writer.WriteLine(base64);
            writer.WriteLine();
        }

        public void writeEnd() {
            writer.WriteLine();
            writer.WriteLine("--" + boundary + "--");
            writer.Close();
        }

    }

    class MhtMerger {

        private List<MhtReader> readers;

        private string boundary;
        private List<QQMessage> messages = new List<QQMessage>();
        private List<string> images = new List<string>();

        public MhtMerger(List<MhtReader> readers) {
            this.readers = readers;
            this.boundary = readers[0].boundary;
        }

        public void Merge() {
            // messages
            foreach (var reader in readers) {
                if (messages.Count == 0) {
                    messages.AddRange(reader.messages);
                } else {
                    int index = 0;
                    foreach (QQMessage msg in reader.messages) {
                        if (checkMessage(messages, index, msg, out index)) {
                            messages.Insert(index, msg);
                        }
                    }
                }
            }
            // images
            var imageNames = new HashSet<string>();
            foreach (MhtReader reader in readers) {
                foreach (var name in reader.imageHashs.Values) {
                    if (imageNames.Add(name)) {
                        this.images.Add(reader.tmpDir + "/" + name);
                    }
                }
            }
        }

        public void WriteToMht(string file) {
            var writer = new MhtWriter(file, boundary);
            writer.writeStart();
            writer.writeHtml(combineToHtml(messages));
            foreach (var imageFile in images) {
                writer.writeImage(imageFile);
            }
            writer.writeEnd();
        }

        public void WriteToSingleHtml(string dir) {
            // write index
            File.WriteAllText(dir + "/index.html", combineToHtml(messages));
            // write images
            Directory.CreateDirectory(dir + "/images");
            foreach (var image in images) {
                File.Copy(image, dir + "/images/" + Path.GetFileName(image));
            }
        }

        public void WriteToMultiHtml(string dir) {
            // group by date
            var groups = messages.GroupBy(msg => msg.date).ToDictionary(gdc => gdc.Key, gdc => gdc.ToList());

            // write index
            string indexContent =
@"<frameset cols='160px, *'>
  <frame src='menu.html' />
  <frame src='about:blank' name='main' />
</frameset>";
            File.WriteAllText(dir + "/index.html", indexContent);

            // left menu
            var menuHtml = new StringBuilder();
            menuHtml.AppendLine(@"<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />
<style>
body { font-size: 12px; line-height: 150%; margin:0px; padding: 10px; }
ul { list-style-type: none; margin: 0px; padding: 0px; }
</style></head>
<body>");
            menuHtml.AppendLine("<ul>");
            foreach (var date in groups.Keys) {
                var dateDisplay = DateTime.ParseExact(date.ToString(), "yyyyMMdd", CultureInfo.CurrentCulture).ToString("yyyy-MM-dd");
                menuHtml.AppendLine(String.Format("<li><a href='{0}.html' target='main'>{1} ({2})</a></li>", date, dateDisplay, groups[date].Count));
            }
            menuHtml.AppendLine("</ul>");
            menuHtml.AppendLine("</body></html>");
            File.WriteAllText(dir + "/menu.html", menuHtml.ToString());

            // date files
            foreach (var date in groups.Keys) {
                File.WriteAllText(dir + "/" + date + ".html", combineToHtml(groups[date]));
            }

            // write images
            Directory.CreateDirectory(dir + "/images");
            foreach (var image in images) {
                File.Copy(image, dir + "/images/" + Path.GetFileName(image));
            }
        }

        private static string combineToHtml(List<QQMessage> messages) {
            var sb = new StringBuilder();
            sb.AppendLine(
            @"<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />
<style>
body { font-size: 12px;}
div { padding: 2px; }
.u { color: #006EFE; }
.c { padding-left: 15px; font-size: 14px;}
.d { font-size: 38px; font-weight: bold; line-height: 55px; border-bottom: 2px solid #CCC; margin-bottom: 10px; }
</style></head>
<body>");
            QQMessage? preMsg = null;
            foreach (var msg in messages) {
                if (preMsg == null || msg.date != preMsg.Value.date) {
                    sb.AppendLine("<div class=d>" + msg.date + "</div>");
                }
                preMsg = msg;
                sb.AppendLine("<div>");
                sb.AppendLine("<div class=u>" + msg.AccountDisplay + " " + msg.TimeDisplay + "</div>");
                sb.AppendLine("<div class=c>" + msg.content + "</div>");
                sb.AppendLine("</div>");
            }
            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private bool checkMessage(List<QQMessage> messages, int fromIndex, QQMessage checkingMsg, out int index) {
            for (index = fromIndex; index < messages.Count; index++) {
                var message = messages[index];
                if (checkingMsg.date < message.date || (checkingMsg.date == message.date && checkingMsg.time < message.time)) {
                    break;
                } else if (checkingMsg.Equals(message)) {
                    return false;
                }
            }
            return true;
        }
    }

    class MhtReader {

        private StreamReader reader;
        private string file;

        // contentLocation -> md5(image).suffix
        public Dictionary<string, string> imageHashs = new Dictionary<string, string>();
        public List<QQMessage> messages = new List<QQMessage>();

        public string boundary;
        public string tmpDir;

        public MhtReader(string file) {
            this.file = file;
            this.reader = new StreamReader(file);
        }

        /// <summary>
        ///  读取并解析mht文件
        /// </summary>
        public void Parse() {
            lookForBoundary();
            createTmpDir();
            parseBlocksAndWriteToTmp();
            parseHtmlToMessages();
            sortMessages();
            fixContentTags();
        }

        /// <summary>
        /// 格式示例（注意boundary一行）：
        /// 
        /// From: <Save by Tencent MsgMgr>
        /// Subject: Tencent IM Message
        /// MIME-Version: 1.0
        /// Content-Type:multipart/related;
        /// 	charset="utf-8"
        /// 	type="text/html";
        /// 	boundary="----=_NextPart_A240C18D_603F_466a_8338.0D1C1796EFC3"
        ///
        /// </summary>
        private void lookForBoundary() {
            do {
                string line = reader.ReadLine().Trim();
                if (line.StartsWith("boundary=")) {
                    boundary = Regex.Match(line, @"boundary=""(.*)""").Groups[1].Value;
                }
            } while (boundary == null && reader.Peek() != -1);

            if (boundary == null) {
                throw new Exception("Invalid mht file, no boundary found: " + file);
            }
        }

        /// <summary>
        /// 创建一个临时文件夹保存解析出来的html和图片
        /// </summary>
        private void createTmpDir() {
            tmpDir = Directory.GetCurrentDirectory() + "/tmp" + DateTime.UtcNow.Ticks;
            Directory.CreateDirectory(tmpDir);
        }

        /// <summary>
        /// 解析mht中的每一块（一个html块加上很多图片块），以文件形式写到临时目录中
        /// </summary>
        private void parseBlocksAndWriteToTmp() {
            string contentType = null;
            string contentTransferEncoding = null;
            string contentLocation = null;
            var content = new StringBuilder();

            var state = State.None;
            while (reader.Peek() != -1) {
                string line = reader.ReadLine().Trim();
                switch (state) {
                    case State.None:
                        if (line == "--" + boundary) state = State.Head;
                        continue;
                    case State.Head:
                        if (line == "") {
                            state = State.Body;
                        } else {
                            // Content-Type:image/jpeg
                            // Content-Transfer-Encoding:base64
                            // Content-Location:{33E8481F-D8C5-42c7-B3E0-FD1484C54196}.dat
                            if (line.StartsWith("Content-Type:")) contentType = line.SubstringAfter(":").Trim();
                            else if (line.StartsWith("Content-Transfer-Encoding:")) contentTransferEncoding = line.SubstringAfter(":").Trim();
                            else if (line.StartsWith("Content-Location:")) contentLocation = line.SubstringAfter(":").Trim();
                        }
                        continue;
                    default /* Body */:
                        // 原mht文件中html部分最后没有空行，所以要判断</table></body></html>
                        if (line == "" || line == "</table></body></html>") {
                            state = State.None;
                            if (contentType == "text/html") {
                                // html块，直接保存到文件
                                string htmlFile = tmpDir + "/index.html";
                                var writer = new StreamWriter(htmlFile);
                                try {
                                    writer.Write(content.ToString());
                                } finally {
                                    writer.Close();
                                }
                            } else if (contentType.StartsWith("image/")) {
                                // 图片块，经过base64转换后，保存到文件
                                var imageContent = content.ToString();
                                var suffix = contentType.SubstringAfter("/"); // png, gif, jpeg
                                if (suffix == "jpeg") {
                                    suffix = "jpg";
                                }
                                var hash = imageContent.Md5();
                                var filename = hash + "." + suffix;

                                // 保存下来，一会儿替换html中的img的src
                                imageHashs.Add(contentLocation, filename);

                                var imageFile = tmpDir + "/" + filename;
                                var binWriter = new BinaryWriter(File.Open(imageFile, FileMode.Create));
                                try {
                                    var bin = System.Convert.FromBase64String(imageContent);
                                    binWriter.Write(bin);
                                } catch (Exception) {
                                    // ignore
                                } finally {
                                    binWriter.Close();
                                }
                            }

                            // 重置几个变量，供下个块使用
                            content = new StringBuilder();
                            contentLocation = null;
                            contentType = null;
                            contentTransferEncoding = null;
                        } else {
                            content.AppendLine(line);
                        }
                        continue;
                }
            }
        }

        /// <summary>
        /// 对html进行解析，得到每个信息的日期、时间、昵称、QQ号、email和内容。使用HtmlAgilityPack库解析html
        /// </summary>
        private void parseHtmlToMessages() {
            string date = null;
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(tmpDir + "/index.html", new System.Text.UTF8Encoding());

            foreach (HtmlAgilityPack.HtmlNode td in doc.DocumentNode.SelectNodes("//td")) {
                // <tr><td style=border-bottom-width:1px;border-bottom-color:#8EC3EB;border-bottom-style:solid;color:#3568BB;font-weight:bold;height:24px;line-height:24px;padding-left:10px;margin-bottom:5px;>日期: 2012/4/11</td></tr>
                var style = td.Attributes["style"];
                if (style != null && style.Value != null && td.InnerText != null && td.InnerText.StartsWith("日期")) {
                    date = td.InnerText.SubstringAfter("日期: ");
                }
                    // <tr><td><div style=color:#006EFE;padding-left:10px;><div style=float:left;margin-right:6px;>大魔头&lt;notyycn@gmail.com&gt;</div>12:28:18</div><div style=padding-left:20px;><font style="font-size:16pt;font-family:'华文楷体','MS Sans Serif',sans-serif;" color='000000'>看字节码干嘛。。。</font></div></td></tr>
                    // <tr><td><div style=color:#006EFE;padding-left:10px;><div style=float:left;margin-right:6px;>519870018(519870018)</div>12:34:43</div><div style=padding-left:20px;><IMG src="{3D48C238-CD47-4d17-9C8F-3593C6D4738B}.dat"></div></td></tr>
                    // <tr><td><div style=color:#42B475;padding-left:10px;><div style=float:left;margin-right:6px;>风自由(23246779)</div>20:11:16</div><div style=padding-left:20px;><font style="font-size:10pt;font-family:'微软雅黑','MS Sans Serif',sans-serif;" color='000000'>就我们两</font></div></td></tr>
                else {
                    var firstDiv = td.FirstChild;
                    string msgStyle = null;
                    if (firstDiv != null) {
                        var styleAttr = firstDiv.Attributes["style"];
                        if (styleAttr != null) {
                            msgStyle = styleAttr.Value;
                        }
                    }
                    if (msgStyle != null && msgStyle.Contains("color") && msgStyle.Contains("padding-left:10px;")) {
                        QQMessage message = new QQMessage();
                        var name = td.FirstChild.FirstChild.InnerText.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&amp;", "&").Replace("&nbsp;", " ");
                        if (name.EndsWith(">")) {
                            message.nickname = name.SubstringBefore("<");
                            message.qqEmail = name.SubstringBetween("<", ">");
                        } else {
                            message.nickname = name.SubstringBefore("(");
                            message.qqNumber = name.SubstringBetween("(", ")");
                        }

                        // 日期：2012/4/11
                        var culture = CultureInfo.CurrentCulture;
                        message.date = Int32.Parse(DateTime.ParseExact(date, new string[] { "yyyy/M/d", "yyyy-M-d", "M/d/yyyy" }, culture, DateTimeStyles.None).ToString("yyyyMMdd"));
                        // 12:28:18
                        message.time = Int32.Parse(DateTime.ParseExact(td.FirstChild.LastChild.InnerText, new string[] { "H:m:s" }, culture, DateTimeStyles.None).ToString("HHmmss"));
                        message.content = td.LastChild.InnerHtml;
                        messages.Add(message);
                    }
                }
            }
        }

        /// <summary>
        /// 有时候时间前后错乱，重新排列
        /// </summary>
        private void sortMessages() {
            messages.Sort((item1, item2) => {
                if (item1.date == item2.date) {
                    return item1.time - item2.time;
                } else {
                    return item1.date - item2.date;
                }
            });
        }

        /// <summary>
        /// 将html中的内容标签进行修正：
        /// 1. img的src改为图片的实际路径（images/md5.图片类型)，以方便比较不同人导出的聊天记录图片是否相同。如果引用的图片不存在，则移除src属性。
        /// 2. 对于font/b/i标签，去掉它们，以统一字体和风格
        /// </summary>
        private void fixContentTags() {
            for (int i = 0; i < messages.Count; i++) {
                QQMessage message = messages[i];
                var doc = new HtmlDocument();
                doc.LoadHtml(message.content);
                var allNodes = doc.DocumentNode.DescendantsAndSelf();
                foreach (var node in allNodes.ToList()) {
                    if (node.Name == "font" || node.Name == "b" || node.Name == "i") {
                        node.ParentNode.AppendChildren(node.ChildNodes);
                        node.ParentNode.RemoveChild(node);
                        continue;
                    }
                    if (node.Name == "img") {
                        var srcAttr = node.Attributes["src"];
                        if (srcAttr != null) {
                            var src = srcAttr.Value;
                            if (imageHashs.ContainsKey(src)) {
                                node.Attributes["src"].Value = "images/" + imageHashs[src];
                            } else {
                                srcAttr.Remove();
                            }
                        }
                    }
                }
                message.content = doc.DocumentNode.OuterHtml;
                messages[i] = message;
            }
        }


        /// <summary>
        /// 删除临时目录
        /// </summary>
        public void DeleteTmpDir() {
            Directory.Delete(this.tmpDir, true);
        }

        /// <summary>
        /// 解析mht文件过程中的三种状态
        /// 
        /// /* None */
        /// /* None */ ------=_NextPart_A240C18D_603F_466a_8338.0D1C1796EFC3
        /// /* Head */ Content-Type: text/html
        /// /* Head */ Content-Transfer-Encoding:7bit
        /// /* Head */ 
        /// /* Body */ <html xmlns="http://www.w3.org/1999/xhtml">
        /// /* Body */ <head>...</head>
        /// /* Body */ <body>...</body>
        /// /* Body */ </html>
        /// /* None */
        /// 
        /// </summary>
        enum State {
            None, Head, Body
        }

    }

    /// <summary>
    /// 表示每一条用户留言
    /// </summary>
    public struct QQMessage {

        /// <summary>
        /// 用户昵称，可能为空
        /// </summary>
        public string nickname;

        /// <summary>
        /// 该用户有QQ号（以圆括号括起来），与qqEmail二选一
        /// </summary>
        public string qqNumber;

        /// <summary>
        /// 该用户没有QQ号，用自己的邮件注册的（以尖括号括起来），与qqNumber二选一
        /// </summary>
        public string qqEmail;

        /// <summary>
        /// 信息日期
        /// </summary>
        public int date;

        /// <summary>
        /// 信息时间
        /// </summary>
        public int time;

        /// <summary>
        /// 信息内容
        /// </summary>
        public string content;

        public bool Equals(QQMessage another) {
            return (qqNumber == another.qqNumber || qqEmail == another.qqEmail) && date == another.date && time == another.time && content == another.content;
        }

        /// <summary>
        /// 按原mht的方式来显示昵称及qq号/Email
        /// </summary>
        public string AccountDisplay {
            get { return nickname + (qqNumber != null ? "(" + qqNumber + ")" : "<" + qqEmail + ">"); }
        }

        /// <summary>
        /// 以易读形式来显示日期时间
        /// </summary>
        public string TimeDisplay {
            get {
                var culture = CultureInfo.CurrentCulture;
                var dateStr = DateTime.ParseExact(date.ToString(), "yyyyMMdd", culture).ToString("yyyy-MM-dd");
                var timeStr = DateTime.ParseExact(time.ToString().PadLeft(6, '0'), "HHmmss", culture).ToString("HH:mm:ss");
                return dateStr + " " + timeStr;
            }
        }
    }
}
