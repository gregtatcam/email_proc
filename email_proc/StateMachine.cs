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
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

namespace email_proc
{
    public delegate void StatusCb(String format, params object [] args);
    public delegate Task DataCb(String mailbox, String message, String msgnum);
    public delegate void ProgressCb(double progress);
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
        ImapConnect connect { get; set; }
        Dictionary<String, String> messageid;
        public StateMachine(String host, int port, String user, String pswd, String dldFile)
        {
            Regex rx = new Regex("^[ ]*([^ \t@]+)(@.+)?$");
            Match m = rx.Match(user);
            if (m.Success)
                this.user = m.Groups[1].Value;
            else
                this.user = user.Trim(' ');
            this.pswd = pswd;
            this.dldFile = dldFile;
            messageid = new Dictionary<string, string>();
            connect = new ImapConnect(host, port);
        }

        ImapCmd CreateImap(Stream strm)
        {
            StreamReader reader = new StreamReader(strm, Encoding.Default, false, 2048, true);
            return new ImapCmd(reader, strm);
        }

        async Task<Dictionary<String,int>> CheckResume(StatusCb status, ProgressCb progress, bool isgmail)
        {
            Dictionary<String, int> downloaded = new Dictionary<string, int>();
            if (File.Exists(dldFile))
            {
                try {
                    Regex rx_mbox = new Regex("^X-Mailbox: (.+)$");
                    Regex rx_msgid = new Regex("^message-id: ([^ ]+)", RegexOptions.IgnoreCase);
                    int read = 0;
                    int cnt = 0;
                    FileInfo info = new FileInfo(dldFile);
                    long length = info.Length;
                    status("Resuming download");
                    status("Calculating resume point...");
                    using (StreamReader reader = new StreamReader(info.OpenRead()))
                    {
                        String line = "";
                        while ((line = await reader.ReadLineAsync()) != null)
                        {
                            Match m = rx_mbox.Match(line);
                            progress((100.0 * cnt) / length);
                            if (m.Success)
                            {
                                if (downloaded.ContainsKey(m.Groups[1].Value) == false)
                                    downloaded[m.Groups[1].Value] = 0;
                                downloaded[m.Groups[1].Value]++;
                                read = 0;
                            }
                            else if (isgmail && read < 5000)
                            {
                                m = rx_msgid.Match(line);
                                if (m.Success)
                                    messageid[m.Groups[1].Value] = "";
                            }
                            read += line.Length;
                            cnt += line.Length + 1;
                        }
                        progress(100);
                    }
                } catch (Exception ex) {; }
            }

            return downloaded;
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
                    cmd = new ImapCmd(reader, connect.stream, true);
                    status("Compression enabled");
                }
                // get the list of mailboxes
                status("Fetching mailbox list...");
                List<String> mailboxes = new List<string>();
                if (await cmd.Run(new ListCommand(async delegate (String mailbox, String ctx)
                        { await Task.Yield(); mailboxes.Add(mailbox); })) != ReturnCode.Ok)
                    throw new FailedException();
                int total = 0;
                int processed = 0;
                bool isgmail = false;
                mailboxes.Sort();
                Regex rx_gmail = new Regex("^\"[[]Gmail[]]/");
                Dictionary<String, int> statusCnt = new Dictionary<string, int>();
                foreach (String mailbox1 in mailboxes)
                {
                    await cmd.Run(new StatusCommand(mailbox1, async delegate (String cnt, String ctx)
                    {
                        await Task.Yield();
                        status("{0} - {1}", mailbox1, cnt);
                        int res = 0;
                        if (int.TryParse(cnt, out res))
                            total += res;
                        statusCnt[mailbox1] = res;
                    }));
                    if (rx_gmail.IsMatch(mailbox1))
                        isgmail = true;
                }
                if (isgmail)
                {
                    mailboxes.Remove("\"[Gmail]/Important\"");
                    mailboxes.Remove("\"[Gmail]/All Mail\"");
                    mailboxes.Remove("\"[Gmail]/Sent Mail\"");
                    mailboxes.Remove("\"[Gmail]/Trash\"");
                    mailboxes.Add("\"[Gmail]/Sent Mail\"");
                    mailboxes.Add("\"[Gmail]/Trash\"");
                    mailboxes.Add("\"[Gmail]/Important\"");
                    mailboxes.Add("\"[Gmail]/All Mail\"");
                }
                status("Total messages: {0}", total);
                Regex rx_msgid = new Regex("^message-id: ([^ \r\n]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                Dictionary<String, int> downloaded = await CheckResume(status, progress, isgmail);
                status("Downloading ...");
                progress(0);
                foreach (String mailbox in mailboxes)
                {
                    int start = downloaded.TryGetValue(mailbox, out start) && start > 0 ? start : 1;
                    int statusMessages = 0;
                    String messages = "0";
                    if (statusCnt.TryGetValue(mailbox, out statusMessages) && start < statusMessages)
                    {
                        if (await cmd.Run(new SelectCommand(mailbox, async delegate (String exists, String ctx)
                                { await Task.Yield(); messages = exists; })) != ReturnCode.Ok)
                            throw new FailedException();
                        status("{0} - {1} messages: ", mailbox, messages);
                        if (messages != "0")
                        {
                            if (mailbox != "\"[Gmail]/Sent Mail\"" && mailbox != "\"[Gmail]/Trash\"" &&
                                mailbox != "\"[Gmail]/Important\"" && mailbox != "\"[Gmail]/All Mail\"")
                            {
                                if (await cmd.Run(new FetchCommand(FetchCommand.Fetch.Body,
                                    async delegate (String message, String msgn)
                                    {
                                        if (isgmail)
                                        {
                                            int len = message.Length > 5000 ? 5000 : message.Length;
                                            Match m = rx_msgid.Match(message, 0, len);
                                            if (m.Success)
                                                messageid[m.Groups[1].Value] = "";
                                        }
                                        progress((100.0 * (processed + int.Parse(msgn))) / total);
                                        await data(mailbox, message, msgn);
                                    }, start)) != ReturnCode.Ok)
                                    status("No messages downloaded");
                            }
                            else
                            {
                                List<String> unique = new List<string>();
                                if (await cmd.Run(new FetchCommand(FetchCommand.Fetch.MessageID,
                                    async delegate (String message, String msgn)
                                    {
                                        await Task.Yield();
                                        Match m = rx_msgid.Match(message);
                                        if (m.Success)
                                        {
                                            String msgid = m.Groups[1].Value;
                                            if (messageid.ContainsKey(msgid))
                                                return;
                                            messageid[msgid] = "";
                                            unique.Add(msgn);
                                        }
                                        else
                                            unique.Add(msgn);
                                    })) != ReturnCode.Ok)
                                {
                                    status("No messages downloaded");
                                }
                                else
                                {
                                    foreach (String n in unique)
                                    {
                                        int num = int.Parse(n);
                                        await cmd.Run(new FetchCommand(FetchCommand.Fetch.Body,
                                            async delegate (String message, String msgn)
                                            {
                                                progress((100.0 * (processed + int.Parse(msgn))) / total);
                                                await data(mailbox, message, msgn);
                                            }, num, num));
                                    }
                                }
                            }
                        }
                    }
                    processed += int.Parse(messages);
                }
                status("Download complete");
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat(@"{0}\email_proc1596.index", Path.GetDirectoryName(dldFile));
                File.Delete(sb.ToString());
            }
            catch { }
            finally
            {
                connect.Close();
            }
        }
    }
}
