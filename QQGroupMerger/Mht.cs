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

        public void merge() {
            // messages
            foreach (var reader in readers) {
                if (messages.Count == 0) {
                    messages.AddRange(reader.messages);
                } else {
                    foreach (QQMessage msg in reader.messages) {
                        int index = 0;
                        if (checkMessage(messages, index - 200, msg, out index)) {
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

        public void writeToMht(string file) {
            var writer = new MhtWriter(file, boundary);
            writer.writeStart();
            writer.writeHtml(combineToHtml());
            foreach (var imageFile in images) {
                writer.writeImage(imageFile);
            }
            writer.writeEnd();
        }

        public void writeToHtml(string file) {
            // write index
            File.WriteAllText(file + "/index.html", combineToHtml());
            // write images
            Directory.CreateDirectory(file + "/images");
            foreach (var image in images) {
                File.Copy(image, file + "/images/" + Path.GetFileName(image));
            }
        }

        private string combineToHtml() {
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
                sb.AppendLine("<div class=u>" + msg.accountDisplay + " " + msg.timeDisplay + "</div>");
                sb.AppendLine("<div class=c>" + msg.content + "</div>");
                sb.AppendLine("</div>");
            }
            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private bool checkMessage(List<QQMessage> messages, int fromIndex, QQMessage checkingMsg, out int index) {
            var i = fromIndex >= 0 ? fromIndex : 0;
            var list = new List<QQMessage>();
            for (; i < messages.Count; i++) {
                var message = messages[i];
                if (checkingMsg.date < message.date || (checkingMsg.date == message.date && checkingMsg.time < message.time)) {
                    break;
                } else if (checkingMsg.date == message.date && checkingMsg.time == message.time) {
                    list.Add(message);
                }
            }
            index = i;
            if (list.Count() == 0) {
                return true;
            }
            var sameUserMessages = (from m in list where m.qqEmail == checkingMsg.qqEmail || m.qqNumber == checkingMsg.qqNumber select m).ToList<QQMessage>();
            if (sameUserMessages.Count == 0) {
                return true;
            }

            for (int j = 0; j < sameUserMessages.Count; j++) {
                var m = sameUserMessages[j];
                if (m.content == checkingMsg.content) return false;
            }

            return true;
        }

    }

    class MhtReader {

        private StreamReader reader;
        private string file;

        // contentLocation -> imagemd5.suffix
        public Dictionary<string, string> imageHashs = new Dictionary<string, string>();
        public List<QQMessage> messages = new List<QQMessage>();

        public string boundary;
        public string tmpDir;

        public MhtReader(string file) {
            this.file = file;
            this.reader = new StreamReader(file);
        }

        public void readAndParse() {
            findBoundary();
            createTmpDir();
            parseBlocksAndWriteToTmp();
            parseHtml();
            sortMessages();
            replaceImageNames();
        }

        private void sortMessages() {
            messages.Sort((item1, item2) => {
                if (item1.date == item2.date) {
                    return item1.time - item2.time;
                } else {
                    return item1.date - item2.date;
                }
            });
        }

        private void replaceImageNames() {
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

        private void parseHtml() {
            string date = null;
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(tmpDir + "/index.html", new System.Text.UTF8Encoding());
            foreach (HtmlAgilityPack.HtmlNode td in doc.DocumentNode.SelectNodes("//td")) {

                // <tr><td style=border-bottom-width:1px;border-bottom-color:#8EC3EB;border-bottom-style:solid;color:#3568BB;font-weight:bold;height:24px;line-height:24px;padding-left:10px;margin-bottom:5px;>日期: 2012/4/11</td></tr>
                var style = td.Attributes["style"];
                if (style != null && style.Value != null && td.InnerText != null && td.InnerText.StartsWith("日期")) {
                    date = substringAfter(td.InnerText, "日期: ");
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
                            message.nickname = substringBefore(name, "<");
                            message.qqEmail = substringBetween(name, "<", ">");
                        } else {
                            message.nickname = substringBefore(name, "(");
                            message.qqNumber = substringBetween(name, "(", ")");
                        }

                        // 日期：2012/4/11
                        message.date = Int32.Parse(DateTime.Parse(date).ToString("yyyyMMdd"));
                        // 12:28:18
                        message.time = Int32.Parse(DateTime.Parse(td.FirstChild.LastChild.InnerText).ToString("HHmmss"));
                        message.content = td.LastChild.InnerHtml;
                        messages.Add(message);
                    }
                }
            }
        }

        private void createTmpDir() {
            tmpDir = Directory.GetCurrentDirectory() + "/tmp" + DateTime.UtcNow.Ticks;
            Directory.CreateDirectory(tmpDir);
        }

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
                        if (line == "") state = State.Body;
                        else {
                            if (line.StartsWith("Content-Type:")) contentType = substringAfter(line, ":").Trim();
                            else if (line.StartsWith("Content-Transfer-Encoding:")) contentTransferEncoding = substringAfter(line, ":").Trim();
                            else if (line.StartsWith("Content-Location:")) contentLocation = substringAfter(line, ":").Trim();
                        }
                        continue;
                    default /* Body */:
                        // 原mht文件中html部分最后没有空行，所以要判断</table></body></html>
                        if (line == "" || line == "</table></body></html>") {
                            state = State.None;
                            if (contentType == "text/html") {
                                string htmlFile = tmpDir + "/index.html";
                                var writer = new StreamWriter(htmlFile);
                                try {
                                    writer.Write(content.ToString());
                                } finally {
                                    writer.Close();
                                }
                            } else if (contentType.StartsWith("image/")) {
                                var imageContent = content.ToString();
                                var suffix = substringAfter(contentType, "/");
                                var hash = md5(imageContent);
                                var filename = hash + "." + suffix;
                                imageHashs.Add(contentLocation, filename);

                                var imageFile = tmpDir + "/" + filename;
                                var binWriter = new BinaryWriter(File.Open(imageFile, FileMode.Create));
                                try {
                                    var bin = System.Convert.FromBase64String(imageContent);
                                    binWriter.Write(bin);
                                } finally {
                                    binWriter.Close();
                                }
                            }
                            content = new StringBuilder();
                            contentLocation = null;
                            contentType = null;
                            contentTransferEncoding = null;
                        } else content.AppendLine(line);
                        continue;
                }
            }
        }

        private static string substringBefore(string str, string key) {
            var index = str.IndexOf(key);
            if (index == -1) return "";
            return str.Substring(0, index);
        }

        private static string substringBetween(string str, string start, string end) {
            var index1 = str.IndexOf(start);
            var index2 = str.LastIndexOf(end);
            if (index1 == -1 || index2 == -1) return "";
            return str.Substring(index1 + start.Length, index2 - index1 - start.Length);

        }

        private static string substringAfter(string str, string key) {
            var index = str.IndexOf(key);
            if (index == -1) return "";
            return str.Substring(index + key.Length);
        }

        private void findBoundary() {
            do {
                string line = reader.ReadLine().Trim();
                if (line.StartsWith("boundary=")) {
                    boundary = Regex.Match(line, @"boundary=""(.*)""").Groups[1].Value;
                }
            } while (boundary == null && reader.Peek() != -1);

            if (boundary == null) {
                throw new Exception("Invalid file, no boundary found: " + file);
            }
        }

        private static string md5(string text) {
            // step 1, calculate MD5 hash from input
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(text);
            byte[] hash = md5.ComputeHash(inputBytes);

            // step 2, convert byte array to hex string
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++) {
                sb.Append(hash[i].ToString("X2"));
            }
            return sb.ToString();
        }

        enum State {
            None, Head, Body
        }

    }

    public struct QQMessage {

        public string nickname;
        public string qqNumber;
        public string qqEmail;
        public int date;
        public int time;
        public string content;

        public string accountDisplay {
            get { return nickname + (qqNumber != null ? "(" + qqNumber + ")" : "<" + qqEmail + ">"); }
        }

        public string timeDisplay {
            get { return date + " " + time; }
        }
    }
}
