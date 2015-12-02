/* 
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
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

namespace email_proc
{
    public delegate void StatusCb(String format, params object [] args);
    public delegate Task DataCb(String mailbox, String message, String msgid, String msgnum);
    public delegate void ProgressCb(double progress);

    public class Mailbox
    {
        public String name;
        public int cnt;
        public int start;
        public Mailbox(String name, int cnt=0, int start=0)
        {
            this.name = name;
            this.cnt = cnt;
            this.start = 1;
        }
    }
    class FailedLoginException : Exception
    {
        public FailedLoginException() { }
    }
    class ConnectionFailedException : Exception
    {
        public ConnectionFailedException() { }
    }
    class FailedException : Exception
    {
        public FailedException() { }
    }
    class SslFailedException : Exception
    {
        public SslFailedException() { }
    }

    class StateMachine
    {
        String user { get; set; }
        String pswd { get; set; }
        String dldFile { get; set; }
        String host { get; set; }
        ImapConnect connect { get; set; }
        Dictionary<String, String> messageid;
        int messagesInAccount { get; set; }
        CancellationToken token { get; set; }
        public StateMachine(CancellationToken token, String host, int port, String user, String pswd, String dldFile)
        {
            Regex rx = new Regex("^[ ]*([^ \t@]+)(@.+)?$");
            Match m = rx.Match(user);
            if (m.Success)
                this.user = m.Groups[1].Value;
            else
                this.user = user.Trim(' ');
            this.pswd = pswd;
            this.dldFile = dldFile;
            this.host = host;
            this.token = token;
            messagesInAccount = 0;
            messageid = new Dictionary<string, string>();
            connect = new ImapConnect(host, port);
        }

        ImapCmd CreateImap(Stream strm)
        {
            StreamReader reader = new StreamReader(strm, Encoding.Default, false, 2048, true);
            return new ImapCmd(token, reader, strm);
        }

        async Task<Dictionary<String,int>> CheckResume(StatusCb status, ProgressCb progress)
        {
            Dictionary<String, int> downloaded = new Dictionary<string, int>();
            String index = Path.Combine(Path.GetDirectoryName(dldFile), "email_proc1596.index");
            if (File.Exists(index))
            {
                try {
                    Regex rx = new Regex("^((\"[^\"]+\")|([^ ]+))(.*)$");
                    int read = 0;
                    FileInfo info = new FileInfo(index);
                    long length = info.Length;
                    bool first = true;
                    using (StreamReader reader = new StreamReader(info.OpenRead()))
                    {
                        String line = await reader.ReadLineAsync(); // host user download file name
                        while ((line = await reader.ReadLineAsync()) != null)
                        {
                            if (first)
                            {
                                first = false;
                                status("Resuming download");
                                status("Calculating resume point...");
                            }
                            Match m = rx.Match(line);
                            progress((100.0 * read) / length);
                            if (m.Success)
                            {
                                if (downloaded.ContainsKey(m.Groups[1].Value) == false)
                                    downloaded[m.Groups[1].Value] = 0;
                                downloaded[m.Groups[1].Value]++;
                                if (m.Groups[4].Success && m.Groups[4].Value != "")
                                    messageid[m.Groups[4].Value] = "";
                            }
                            read += line.Length + 1;
                        }
                        if (downloaded.Count > 0)
                            progress(100);
                    }
                } catch (Exception ex) {; }
            }

            return downloaded;
        }

        bool GmailSpecial(String name)
        {
            return name == "\"[Gmail]/Sent Mail\"" || name == "\"[Gmail]/Trash\"" ||
                            name == "\"[Gmail]/Important\"" || name == "\"[Gmail]/All Mail\""; 
        }
        async Task<List<Mailbox>> GetMailboxesList(ImapCmd cmd, StatusCb status, ProgressCb progress)
        {
            // get the list of mailboxes
            status("Fetching mailbox list...");
            List<Mailbox> mailboxes = new List<Mailbox>();
            if (await cmd.Run(new ListCommand(async delegate (String mailbox, String ctx)
            { await Task.Yield(); mailboxes.Add(new Mailbox(mailbox)); })) != ReturnCode.Ok)
                throw new FailedException();
            mailboxes.Sort((Mailbox m1, Mailbox m2) => {
                bool g1 = GmailSpecial(m1.name);
                bool g2 = GmailSpecial(m2.name);
                if (g1 && g2)
                    return String.Compare(m2.name, m1.name); // reverse
                else if (g1)
                    return 1;
                else if (g2)
                    return -1;
                else
                    return String.Compare(m1.name, m2.name);
            });
            for (int i = 0; i < mailboxes.Count; i++)
            {
                await cmd.Run(new StatusCommand(mailboxes[i].name, async delegate (String cnt, String ctx)
                {
                    await Task.Yield();
                    status("{0} - {1}", mailboxes[i].name, cnt);
                    mailboxes[i].cnt = int.Parse(cnt);
                    messagesInAccount += int.Parse(cnt);
                }));
            }
            status("Total messages: {0}", messagesInAccount);
            Dictionary<String, int> downloaded = await CheckResume(status, progress);

            foreach (String mbox in downloaded.Keys)
            {
                mailboxes.ForEach((Mailbox m) => {
                    if (m.name == mbox)
                        m.start = downloaded[mbox] > 0 ? downloaded[mbox] : 1;
                });
            }

            return mailboxes;
        }

        public async Task Start(StatusCb status, DataCb data, ProgressCb progress, bool compression=false)
        {
            try
            {
                status("Connecting...");
                if (await connect.Connect() != ConnectStatus.Ok)
                    throw new ConnectionFailedException();
                ImapCmd cmd = null;
                if (await connect.ConnectTls() == ConnectStatus.Failed)
                {
                    cmd = CreateImap(connect.stream);
                    if (await cmd.Run(new StarttlsCommand()) != ReturnCode.Ok)
                        throw new SslFailedException();
                }
                cmd = CreateImap(connect.stream);

                // initial Capability
                await cmd.Run(new NullCommand());

                // login to the server
                if (await cmd.Run(new LoginCommand(user, pswd)) != ReturnCode.Ok)
                    throw new FailedLoginException();
                status("Logged in to the email server");
                // check if compression is supported
                if (compression && await cmd.Run(new CapabilityCommand(new String[] { "compress=deflate" })) == ReturnCode.Ok &&
                    await cmd.Run(new DeflateCommand()) == ReturnCode.Ok)
                {
                    StreamReader reader = new StreamReader(new DeflateStream(connect.stream, CompressionMode.Decompress, true));
                    cmd = new ImapCmd(token, reader, connect.stream, true);
                    status("Compression enabled");
                }
                // get the list of mailboxes
                status("Fetching mailbox list...");
                List<Mailbox> mailboxes = await GetMailboxesList(cmd,status,progress);
                status("Downloading ...");
                progress(0);
                int processed = 0;
                Regex rx_msgid = new Regex("^message-id: ([^ \r\n]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                Func<String, String> getMsgId = (String message) =>
                {
                    int len = message.Length > 5000 ? 5000 : message.Length;
                    Match m = rx_msgid.Match(message, 0, len);
                    return (m.Success ? m.Groups[1].Value : "");
                };
                foreach (Mailbox mailbox in mailboxes)
                {
                    if (token.IsCancellationRequested)
                        return;
                    status("{0} - {1} messages: ", mailbox.name, mailbox.cnt);
                    if (mailbox.start < mailbox.cnt)
                    {
                        if (await cmd.Run(new SelectCommand(mailbox.name)) != ReturnCode.Ok)
                            throw new FailedException();

                        if (!GmailSpecial(mailbox.name))
                        {
                            if (await cmd.Run(new FetchCommand(FetchCommand.Fetch.Body,
                                async delegate (String message, String msgn)
                                {
                                    String msgid = "";
                                    if (host == "imap.gmail.com")
                                    {
                                        msgid = getMsgId(message);
                                        messageid[msgid] = "";
                                    }
                                    progress((100.0 * (processed + int.Parse(msgn))) / messagesInAccount);
                                    await data(mailbox.name, message, msgid, msgn);
                                }, mailbox.start)) != ReturnCode.Ok)
                                status("No messages downloaded");
                        }
                        else
                        {
                            List<String> unique = new List<string>();
                            progress(0);
                            if (await cmd.Run(new FetchCommand(FetchCommand.Fetch.MessageID,
                                async delegate (String message, String msgn)
                                {
                                    String msgid = getMsgId(message);
                                    if (msgid != "" && messageid.ContainsKey(msgid))
                                        return;
                                    else
                                    {
                                        messageid[msgid] = "";
                                        unique.Add(msgn);
                                    }
                                    progress(100.0 * int.Parse(msgn) / mailbox.cnt);
                                    await Task.Yield();
                                })) != ReturnCode.Ok)
                            {
                                progress(0);
                                status("No messages downloaded");
                            }
                            else
                            {
                                progress(0);
                                status("Downloading {0} out of {1}", unique.Count, mailbox.cnt);
                                foreach (String n in unique)
                                {
                                    int num = int.Parse(n);
                                    await cmd.Run(new FetchCommand(FetchCommand.Fetch.Body,
                                        async delegate (String message, String msgn)
                                        {
                                            progress((100.0 * (processed + int.Parse(msgn))) / messagesInAccount);
                                            await data(mailbox.name, message, getMsgId(message), msgn);
                                        }, num, num));
                                }
                            }
                        }
                    }
                    processed += mailbox.cnt;
                }
                status("Download complete");
            }
            catch { }
            finally
            {
                connect.Close();
            }
        }
    }
}
