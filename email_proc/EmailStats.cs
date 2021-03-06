﻿/* 
 * Copyright(c) 2015-2016 Gregory Tsipenyuk<gt303@cam.ac.uk>
 *  
 * Permission to use, copy, modify, and distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above 
 * copyright notice and this permission notice appear in all copies. 
 *  
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES 
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF 
 * MERCHANTABILITY AND FITNESS.IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES 
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace email_proc
{
    class EmailStats
    {
        public delegate void StatusCb(bool cr, string crcondition, string format, params object[] args);
        public delegate void ProgressCb(double progress);
        Dictionary<String, int> mailboxes { get; set; }
        Dictionary<String, String> messageid { get; set; }
        CancellationToken token { get; set; }

        void GetUniqueMailbox(String m)
        {
            GetUnique(mailboxes, m.Trim('\"').ToLowerInvariant());
        }

        void InitMailboxes()
        {
            GetUniqueMailbox("inbox");
            GetUniqueMailbox("sent");
            GetUniqueMailbox("sent messages");
            GetUniqueMailbox("\"[Gmail]/Sent Mail\"");
            GetUniqueMailbox("trash");
            GetUniqueMailbox("\"[Gmail]/Trash\"");
            GetUniqueMailbox("junk");
            GetUniqueMailbox("deleted");
            GetUniqueMailbox("deleted messages");
            GetUniqueMailbox( "spam");
            GetUniqueMailbox( "\"[Gmail]/Spam\"");
            GetUniqueMailbox( "\"[Gmail]/All Mail\"");
            GetUniqueMailbox("\"[Gmail]/Important\"");
            GetUniqueMailbox( "drafts");
            GetUniqueMailbox( "\"[Gmail]/Drafts\"");
        }
        public EmailStats(CancellationToken token)
        {
            mailboxes = new Dictionary<string, int>();
            messageid = new Dictionary<string, String>();
            this.token = token;
            InitMailboxes();
        }

        String GetUnique(Dictionary<String,int> dict, String key)
        {
            if (key == "")
                return "";
            if (dict.ContainsKey(key) == false)
                dict[key] = dict.Count + 1;
            return String.Format("{0}", dict[key]);
        }

        String GetMessageId(String key)
        {
            if (key == "")
                return "";
            if (messageid.ContainsKey(key))
                return messageid[key];
            String sha = Sha1(key);
            messageid[key] = sha;
            return sha;
        }

        String GetSubject(String subject)
        {
            if (subject == "")
                return "";
            subject = Regex.Replace(subject, "[\r\n\t ]+", " ");
            Regex re = new Regex("(re|fw|fwd): (.*)$", RegexOptions.IgnoreCase);
            Match m = re.Match(subject);
            if (m.Success)
            {
                subject = m.Groups[2].Value;
                return ("re/fw: " + Sha1(subject));
            }
            return Sha1(subject);
        }

        String GetInReplyTo(String str)
        {
            if (str == "")
                return str;
            String[] ids = Regex.Split(str, "[ \r\n]");
            StringBuilder sb = new StringBuilder();
            String sha1 = "";
            foreach (String msgid in ids)
            {
                if (messageid.ContainsKey(msgid))
                    sha1 = messageid[msgid];
                else
                    sha1 = Sha1(msgid);
                if (sb.Length == 0)
                    sb.Append(sha1);
                else
                    sb.AppendFormat(",{0}", sha1);
            }
            return sb.ToString();
        }

        String GetMailbox(String mailbox)
        {
            if (mailbox == "")
                return GetUnique(mailboxes, "inbox");
            mailbox = mailbox.ToLower().Trim('\"');
            string[] labels = Regex.Split(mailbox, ",");
            StringBuilder sb = new StringBuilder();
            foreach (string label in labels)
            {
                if (Regex.IsMatch(label, "[[]Gmail[]]", RegexOptions.IgnoreCase))
                {
                    if (sb.Length == 0)
                        sb.Append(GetUnique(mailboxes, label));
                    else
                        sb.AppendFormat(",{0}", GetUnique(mailboxes, label));
                }
                else
                {
                    string[] parts = Regex.Split(mailbox, "/");
                    StringBuilder sb1 = new StringBuilder();
                    foreach (string part in parts)
                    {
                        if (sb1.Length == 0)
                            sb1.Append(GetUnique(mailboxes, part));
                        else
                            sb1.AppendFormat("/{0}", GetUnique(mailboxes, part));
                    }
                    if (sb.Length == 0)
                        sb.Append(sb1);
                    else
                        sb.AppendFormat(",{0}", sb1);
                }
            }
            return sb.ToString();
        }

        String MatchEmail(String re_str, String addr)
        {
            Regex re = new Regex(re_str);
            Match m = re.Match(addr);
            if (m.Success)
                return m.Groups[1].Value.ToLowerInvariant();
            else
                return null;
        }

        String GetAddrList(String str)
        {
            String[] addrs = Regex.Split(str, "[,;]");
            StringBuilder sb = new StringBuilder();
            int c = 0;
            foreach(String addr in addrs)
            {
                if (c > 0)
                    sb.AppendFormat(",{0}", Sha1(EmailAddr(addr)));
                else
                    sb.Append(Sha1(EmailAddr(addr)));
                c++;
            }
            return sb.ToString();
        }

        String EmailAddr(String addr)
        {
            // match the rightmost email address, so in "xxx1@foo.com" <xxx2@foo.com>, xxx2@foo.com is matched
            String re_str = "([^\'\"\t\r\n <@:[]+@[^]\t :<>\'\"\r\n]+)";
            String matched = MatchEmail("<" + re_str + ">[ \t]*$", addr);
            if (matched == null && (matched = MatchEmail("<" + re_str + ">", addr)) == null && (matched = MatchEmail(re_str, addr)) == null)
                return "";
            else
                return matched.ToLowerInvariant();
        }

        bool MessageidUnique(String msgid)
        {
            return msgid == "" || messageid.ContainsKey(msgid) == false;
        }

        Task WriteStatsLine(StreamWriter writer, String format, params object [] args)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(format, args);
            return writer.WriteLineAsync(sb.ToString());
        }

        async Task<int> CompressedSize(byte[] buff)
        {
            MemoryStream deflated = new MemoryStream();
            MemoryStream original = new MemoryStream(buff);
            using (DeflateStream deflate = new DeflateStream(deflated, CompressionLevel.Optimal,true))
            {
                await original.CopyToAsync(deflate);
            }
            return (int)deflated.Length;
        }

        String Sha1(byte[] buff)
        {
            HashAlgorithm sha1 = SHA1.Create();
            String sha = Convert.ToBase64String(sha1.ComputeHash(buff));
            return (Regex.Replace(sha, "/", "o057"));
        }

        String Sha1(String buff)
        {
            if (buff == "")
                return "";
            return Sha1(Encoding.Default.GetBytes(buff));
        }

        async Task TraverseEmail(StreamWriter filew, int id, int part, Email email)
        {
            await WriteStatsLine(filew, "part: {0}", part);
            int csize = await CompressedSize(email.headers.GetBytes());
            await WriteStatsLine(filew, "headers: {0} {1} {2}", email.headers.GetNumberOfHeaders(), email.headers.size, csize);
            await WriteStatsLine(filew, "contenttype: {0}", email.headers.contentTypeFullStr);
            switch (email.content.dataType)
            {
                case DataType.Data:
                    if (email.headers.contentType == ContentType.Audio || email.headers.contentType == ContentType.Video ||
                        email.headers.contentType == ContentType.Image || email.headers.contentType == ContentType.Application)
                    {
                        byte[] buff = email.content.GetBytes();
                        csize = await CompressedSize(buff);
                        String sha1 = Sha1(buff);
                        await WriteStatsLine(filew, "attachment: {0} {1} {2}", sha1, email.content.size, csize);
                    }
                    else
                    {
                        csize = await CompressedSize(email.content.GetBytes());
                        await WriteStatsLine(filew, "body: {0} {1}", email.content.size, csize);
                    }
                    break;
                case DataType.Message:
                    await WriteStatsLine(filew, "start rfc822: {0}", id);
                    await TraverseEmail(filew, (id + 1), 0, email.content.data[0]);
                    await WriteStatsLine(filew, "end rfc822: {0}", id);
                    break;
                case DataType.Multipart:
                    await WriteStatsLine(filew, "start multipart {0} {1}:", id, email.content.data.Count);
                    int p = 0;
                    foreach (Email e in email.content.data)
                        await TraverseEmail(filew, id + 1, p++, e);
                    await WriteStatsLine(filew, "end multipart {0}", id);
                    break;
            }
        }

        public async Task Start(String dir, String file, StatusCb status, ProgressCb pcb)
        {
            StreamWriter filew = null;
            try {
                status(false, "", "Started statistics processing.");
                MessageReader reader = new MessageReader(file);
                long size = reader.BaseStream.Length;
                double progress = .0;
                StringBuilder sb = new StringBuilder();
                sb = new StringBuilder();
                sb.AppendFormat(@"{0}\stats{1}.out", dir, DateTime.Now.ToFileTime());
                file = sb.ToString();
                filew = new StreamWriter(file);
                DateTime start_time = DateTime.Now;
                await WriteStatsLine(filew, "archive size: {0}\n", size);
                int count = 1;
                await EmailParser.ParseMessages(token, reader, async delegate (Message message, Exception reason)
                {
                    status(true, "^message:", "message: {0}", count++);
                    try
                    {
                        if (null == message)
                        {
                            status(false, "", "message parsing failed: " + (null != reason ? reason.Message : ""));
                            await WriteStatsLine(filew, "--> start");
                            await WriteStatsLine(filew, "<-- end failed to process: {0}", (null != reason) ? reason.Message : "");
                            return;
                        }
                        // display progress
                        progress += message.size;
                        double pct = (100.0 * progress / (double)size);
                        pcb(pct);
                        // get required headers
                        Dictionary<String, String> headers = message.email.headers.GetDictionary(new Dictionary<string, string>()
                        { {"from","" }, { "cc", "" }, {"subject","" }, {"date","" },
                        { "to",""}, {"bcc", "" }, { "in-reply-to","" }, {"reply-to","" }, {"content-type","" }, {"message-id","" }, { "x-gmail-labels",""}});
                        String msgid = headers["message-id"];
                        // get unique messages
                        if (msgid != null && msgid != "" && MessageidUnique(msgid) == false)
                        {
                            return;
                        }
                        await WriteStatsLine(filew, "--> start");
                        int csize = await CompressedSize(message.GetBytes());
                        await WriteStatsLine(filew, "Full Message: {0} {1}", message.size, csize);
                        await WriteStatsLine(filew, "Hdrs");
                        await WriteStatsLine(filew, "from: {0}", Sha1(EmailAddr(headers["from"])));
                        await WriteStatsLine(filew, "to: {0}", GetAddrList(headers["to"]));
                        await WriteStatsLine(filew, "cc: {0}", GetAddrList(headers["cc"]));
                        await WriteStatsLine(filew, "bcc: {0}", GetAddrList(headers["bcc"]));
                        await WriteStatsLine(filew, "date: {0}", headers["date"]);
                        await WriteStatsLine(filew, "subject: {0}", GetSubject(headers["subject"]));
                        await WriteStatsLine(filew, "mailbox: {0}", GetMailbox(headers["x-gmail-labels"]));
                        await WriteStatsLine(filew, "messageid: {0}", GetMessageId(headers["message-id"]));
                        await WriteStatsLine(filew, "inreplyto: {0}", GetInReplyTo(headers["in-reply-to"]));
                        await WriteStatsLine(filew, "replyto: {0}", GetAddrList(headers["reply-to"]));
                        await WriteStatsLine(filew, "Parts:");
                        await TraverseEmail(filew, 0, 0, message.email);
                        await WriteStatsLine(filew, "<-- end");
                    }
                    catch (Exception ex)
                    {
                        await WriteStatsLine(filew, "<-- end failed to process: {0}, {1}", ex.Message, ex.StackTrace);
                    }
                });
                status(false, "", "Statistics is generated in file {0}", file);
                TimeSpan span = DateTime.Now - start_time;
                status(false, "", "Processing time: {0} seconds", span.TotalSeconds);
            } 
            catch (Exception ex)
            {
                status(false, "", "Statistics failed: {0}", ex.Message);
            }
            finally
            {
                if (filew != null)
                    filew.Close();
            }
        }
    }
}
