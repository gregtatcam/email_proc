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
    public delegate void StatusCb(String format, params string [] args);
    public delegate Task DataCb(String mailbox, String message, String msgnum);
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
        ImapConnect connect { get; set; }
        public StateMachine(String host, int port, String user, String pswd)
        {
            Regex rx = new Regex("^[ ]*([^ \t@]+)(@.+)?$");
            Match m = rx.Match(user);
            if (m.Success)
                this.user = m.Groups[1].Value;
            else
                this.user = user.Trim(' ');
            this.pswd = pswd;
            connect = new ImapConnect(host, port);
        }

        ImapCmd CreateImap(Stream strm)
        {
            StreamReader reader = new StreamReader(strm, Encoding.Default, false, 2048, true);
            return new ImapCmd(reader, strm);
        }
        public async Task Start(StatusCb status, DataCb data, bool compression=false)
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
                List<String> mailboxes = new List<string>();
                if (await cmd.Run(new ListCommand(async delegate (String mailbox, String ctx) 
                        { await Task.Yield(); mailboxes.Add(mailbox); })) != ReturnCode.Ok)
                    throw new FailedException();
                status("Downloading ...");
                foreach (String mailbox in mailboxes)
                {
                    String messages = "0";
                    if (await cmd.Run(new SelectCommand(mailbox, async delegate (String exists, String ctx) 
                            { await Task.Yield(); messages = exists; })) != ReturnCode.Ok)
                        throw new FailedException();
                    status("{0} - {1} messages: ", mailbox, messages);
                    if (messages != "0" &&
                        await cmd.Run(new FetchCommand(FetchCommand.Fetch.Body, 
                            async delegate (String message, String msgn) { await data(mailbox, message, msgn); })) != ReturnCode.Ok)
                            status("No messages downloaded");
  
                }
                status("Download complete");
            }
            catch (FailedLoginException)
            {
                status("Logged in failed, invalid user or password");
            }
            catch (FailedException)
            {
                status("Failed to download");
            }
            catch (SslFailedException)
            {
                status("Failed to establish secure connection");
            }
            finally
            {
                connect.Close();
            }
        }
    }
}
