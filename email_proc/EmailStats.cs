using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace email_proc
{
    class EmailStats
    {
        public delegate void StatusCb(bool cr, string format, params object[] args);
        public delegate void ProgressCb(double progress);
        Dictionary<String, int> emailAddr { get; set; }
        Dictionary<String, int> attachments { get; set; }
        Dictionary<String, int> mailboxes { get; set; }
        Dictionary<String, int> subject { get; set; }
        Dictionary<String, int> messageid { get; set; }
        Dictionary<String, int> inreplyto { get; set; }

        void InitMailboxes()
        {
            GetUnique(mailboxes, "inbox");
            GetUnique(mailboxes, "sent");
            GetUnique(mailboxes, "sent messages");
            GetUnique(mailboxes, "\"[Gmail]/Sent Mail\"");
            GetUnique(mailboxes, "trash");
            GetUnique(mailboxes, "\"[Gmail]/Trash\"");
            GetUnique(mailboxes, "junk");
            GetUnique(mailboxes, "deleted");
            GetUnique(mailboxes, "deleted messages");
            GetUnique(mailboxes, "spam");
            GetUnique(mailboxes, "\"[Gmail]/Spam\"");
            GetUnique(mailboxes, "\"[Gmail]/All Mail\"");
            GetUnique(mailboxes, "\"[Gmail]/Important\"");
            GetUnique(mailboxes, "drafts");
            GetUnique(mailboxes, "\"[Gmail]/Drafts\"");
        }
        public EmailStats()
        {
            emailAddr = new Dictionary<string, int>();
            attachments = new Dictionary<string, int>();
            mailboxes = new Dictionary<string, int>();
            subject = new Dictionary<string, int>();
            messageid = new Dictionary<string, int>();
            inreplyto = new Dictionary<string, int>();
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

        String GetInReplyTo(String str)
        {
            String[] ids = Regex.Split(str, "[ \r\n]");
            StringBuilder sb = new StringBuilder();
            foreach (String msgid in ids)
            {
                if (messageid.ContainsKey(msgid) == false)
                    inreplyto[msgid] = 0;
                String u = GetUnique(messageid, msgid);
                if (sb.Length == 0)
                    sb.Append(u);
                else
                    sb.AppendFormat(",{0}", u);
            }
            return sb.ToString();
        }

        bool MessageidUnique(String msgid)
        {
            bool res = msgid == "" || messageid.ContainsKey(msgid) == false || inreplyto.ContainsKey(msgid);
            if (inreplyto.ContainsKey(msgid))
                inreplyto.Remove(msgid);
            return res;
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

        async Task<String> Sha1(byte[] buff)
        {
            HashAlgorithm sha1 = SHA1.Create();
            return await Task.Run<String>(() => { return Convert.ToBase64String(sha1.ComputeHash(buff)); });
        }

        async Task TraverseEmail(StreamWriter filew, int id, int part, Email email)
        {
            await WriteStatsLine(filew, "part: {0}", part);
            int csize = await CompressedSize(email.headers.GetBytes());
            await WriteStatsLine(filew, "headers: {0} {1} {2}", email.headers.lines, email.headers.size, csize);
            await filew.WriteLineAsync("H---------------------------");
            await filew.WriteAsync(email.headers.GetString());
            await filew.WriteLineAsync("----------------------------");
            switch (email.content.dataType)
            {
                case DataType.Data:
                    await filew.WriteLineAsync("C---------------------------");
                    await filew.WriteAsync(email.content.GetString());
                    await filew.WriteLineAsync("----------------------------");
                    if (email.headers.contentType == ContentType.Audio || email.headers.contentType == ContentType.Video ||
                        email.headers.contentType == ContentType.Image || email.headers.contentType == ContentType.Application)
                    {
                        byte[] buff = email.content.GetBytes();
                        csize = await CompressedSize(buff);
                        String sha1 = await Sha1(buff);
                        await WriteStatsLine(filew, "attachment: {0} {1} {2}", GetUnique(attachments, sha1), 
                            email.content.size, csize);
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
                status(false, "Started statistics processing.");
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
                await EmailParser.ParseMessages(reader, async delegate (Message message)
                {
                    try {
                        // display progress
                        progress += message.size;
                        double pct = (100.0 * progress / (double)size);
                        pcb(pct);
                        // get required headers
                        Dictionary<String, String> headers = message.email.headers.GetDictionary(new Dictionary<string, string>()
                        { {"from","" }, { "cc", "" }, {"subject","" }, {"x-mailbox","" }, {"date","" },
                        { "to",""}, { "in-reply-to","" }, {"content-type","" }, {"message-id","" } });
                        String msgid = headers["message-id"];
                        // get unique messages
                        if (msgid != null && msgid != "" && MessageidUnique(msgid) == false)
                            return;
                        await WriteStatsLine(filew, "--> start");
                        int csize = await CompressedSize(message.GetBytes());
                        await WriteStatsLine(filew, "Full Message: {0} {1}", message.size, csize);
                        await WriteStatsLine(filew, "Hdrs");
                        await WriteStatsLine(filew, "from: {0}", GetUnique(emailAddr, headers["from"]));
                        await WriteStatsLine(filew, "to: {0}", GetUnique(emailAddr, headers["to"]));
                        await WriteStatsLine(filew, "cc: {0}", GetUnique(emailAddr, headers["cc"]));
                        await WriteStatsLine(filew, "date: {0}", headers["date"]);
                        await WriteStatsLine(filew, "subject: {0}", GetUnique(subject, headers["subject"]));
                        await WriteStatsLine(filew, "mailbox: {0}", GetUnique(mailboxes, headers["x-mailbox"].ToLower()));
                        await WriteStatsLine(filew, "messageid: {0}", GetUnique(messageid, headers["message-id"]));
                        await WriteStatsLine(filew, "inreplyto: {0}", GetInReplyTo(headers["in-reply-to"]));
                        await WriteStatsLine(filew, "Parts");
                        await TraverseEmail(filew, 0, 0, message.email);
                        await WriteStatsLine(filew, "<-- end");
                    }
                    catch (Exception ex)
                    {
                        await WriteStatsLine(filew, "<-- end failed to process: {0}", ex.Message);
                    }
                });
                status(false, "Statistics is generated in file {0}", file);
                TimeSpan span = DateTime.Now - start_time;
                status(false, "Processing time: {0} seconds", span.Seconds);
            } 
            catch (Exception ex)
            {
                status(false, "Statistics failed: {0}", ex.Message);
            }
            finally
            {
                if (filew != null)
                    filew.Close();
            }
        }
    }
}
