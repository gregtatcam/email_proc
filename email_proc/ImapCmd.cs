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
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.IO.Compression;

namespace email_proc
{
    public class CancelException : Exception
    {
        public CancelException() : base() { }
    }

    public class ServerIOException : Exception
    {
        public ServerIOException(String msg) : base(msg) { }
    }

    public enum ReturnCode { Ok, No, Bad, Failed, Cancelled }

    public abstract class Command
    {
        public delegate Task CmdCb(String str, String str1 = null);
        public enum MoreInput { Line, Chunk, Done }
        public enum InputType { Line, Chunk }
        static int tag = 0;
        StringBuilder sb;
        protected CmdCb cb;
        public String Cmd { get { return sb.ToString(); } }
        int Tag { get { return tag; } }
        internal ReturnCode Ret { get; set; }
        public class Return
        {
            public Return(MoreInput r, int ?s=null)
            {
                Code = r;
                Size = s;
            }
            public MoreInput Code { get; protected set; }
            public int ?Size { get; protected set; }
        }

        public Command(CmdCb cb=null, bool isnull=false)
        {
            sb = new StringBuilder();
            if (isnull == false)
                sb.AppendFormat("tag{0:d6} ", ++tag);
            this.cb = cb;
        }
        protected void Append(string format, params object[] args)
        {
            sb.AppendFormat(format, args);
        }
        protected virtual async Task<Return> ParseInternal(InputType t, String str)
        {
            await Task.Yield();
            return (new Return(MoreInput.Line));
        }
        public async Task<Return> Parse(InputType t, String str)
        {
            if (Done(str))
                return (new Return(MoreInput.Done));
            else
                return await ParseInternal(t, str);
        }
        public bool Done(String str)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("^tag{0:d6} (ok|no|bad)", tag);
            Regex re = new Regex(sb.ToString(), RegexOptions.IgnoreCase);
            Match m = re.Match(str);
            if (m.Success)
            {
                switch (m.Groups[1].Value.ToLower())
                {
                    case "ok":
                        Ret = ReturnCode.Ok;
                        break;
                    case "no":
                        Ret = ReturnCode.No;
                        break;
                    case "bad":
                        Ret = ReturnCode.Bad;
                        break;
                    default:
                        Ret = ReturnCode.Failed;
                        break;
                }
                return true;
            }
            else
                return false;
        }
    }

    /*
        * LIST (\HasChildren \Noselect) "/" "[Gmail]"
        * LIST (\All \HasNoChildren) "/" "[Gmail]/All Mail"
    */
    public class ListCommand : Command {
        public ListCommand(CmdCb cb) : base(cb) {
            Append("list \"\" *\r\n");
        }
        protected override async Task<Return> ParseInternal(InputType t, string str)
        {
            Regex re = new Regex(@"^\* list \(([^)]+)\) ([^ ]+) (.+)$", RegexOptions.IgnoreCase);
            Match m = re.Match(str);
            if (m.Success && Regex.IsMatch(m.Groups[1].Value, "noselect", RegexOptions.IgnoreCase) == false)
            {
                await cb(m.Groups[3].Value);
            }
            return (new Return(MoreInput.Line));
        }
    }

    /*
        * CAPABILITY IMAP4rev1 UNSELECT IDLE NAMESPACE QUOTA ID XLIST CHILDREN X-GM-EXT-1 UIDPLUS COMPRESS=DEFLATE ENABLE MOVE CONDSTORE 
            ESEARCH UTF8=ACCEPT LIST-EXTENDED LIST-STATUS
    */
    public class CapabilityCommand : Command
    {
        String[] reqCapabilities { get; set; }
        public CapabilityCommand(CmdCb cb) : base(cb)
        {
            this.reqCapabilities = null;
            Append("capability\r\n");
        }
        public CapabilityCommand(String [] reqCapabilities) : base(null)
        {
            this.reqCapabilities = reqCapabilities;
            Append("capability\r\n");
        }

        protected override async Task<Return> ParseInternal(InputType t, string str)
        {
            Regex re = new Regex(@"^\* \(?(.+)\)?$");
            Match m = re.Match(str);
            if (m.Success)
            {
                if (cb != null)
                    await cb(m.Groups[1].Value);
                else
                {
                    String[] caps = Regex.Split(m.Groups[1].Value, " ");
                    if (caps.Length == 0)
                        Ret = ReturnCode.No;
                    else
                    {
                        StringComparer comp = StringComparer.Create(System.Globalization.CultureInfo.CurrentCulture, true);
                        for (int i = 0; i < reqCapabilities.Length; i++)
                        {
                            if (caps.Contains<String>(reqCapabilities[i],comp) == false)
                            {
                                Ret = ReturnCode.No;
                                break;
                            }
                        }
                        Ret = ReturnCode.Ok;
                    }
                }
            }
            return (new Return(MoreInput.Line));
        }
    }

    /*
        * STATUS "inbox" (MESSAGES 11)
    */
    public class StatusCommand : Command
    {
        public StatusCommand(String mailbox, CmdCb cb) : base(cb)
        {
            Append("status {0} (messages)\r\n", mailbox);
        }

        protected override async Task<Return> ParseInternal(InputType t, string str)
        {
            Regex re = new Regex(@"^\* status [^)]+ \(messages ([0-9]+)\)", RegexOptions.IgnoreCase);
            Match m = re.Match(str);
            if (m.Success)
                await cb(m.Groups[1].Value);
            else
                await cb("0");
            return (new Return(MoreInput.Line));
        }
    }

    public class LoginCommand : Command
    {
        public LoginCommand(String user, String password) : base(null)
        {
            Append("login {0} {1}\r\n", user, password);
        }
    }

    /*
        * 11 EXISTS
     */
    public class SelectCommand : Command
    {
        public SelectCommand(String mailbox,CmdCb cb=null) : base(cb)
        {
            Append("select {0}\r\n", mailbox);
        }

        protected override async Task<Return> ParseInternal(InputType t, string str)
        {
            Regex re = new Regex(@"^\* ([0-9]+) exists", RegexOptions.IgnoreCase);
            Match m = re.Match(str);
            if (m.Success && cb != null)
                await cb(m.Groups[1].Value);
            return (new Return(MoreInput.Line));
        }
    }

    public class DeflateCommand : Command
    {
        public DeflateCommand() : base(null)
        {
            Append("compress deflate\r\n");
        }
    }

    public class StarttlsCommand : Command
    {
        public StarttlsCommand() : base(null)
        {
            Append("starttls\r\n");
        }
    }

    public class FetchCommand : Command
    {
        public enum Fetch { Header, Text, Body, MessageID}
        String msgnum { get; set; }
        public FetchCommand(Fetch fetch, CmdCb cb, int? start=null, int? end= null) : base(cb)
        {
            if (start == null)
                Append("fetch 1:* ");
            else if (end == null)
                Append("fetch {0}:* ", start);
            else 
                Append("fetch {0}:{1} ", start, end);
            switch (fetch)
            {
                case Fetch.Header:
                    Append("body.peek[header]");
                    break;
                case Fetch.Body:
                    Append("body.peek[]");
                    break;
                case Fetch.Text:
                    Append("body.peek[text]");
                    break;
                case Fetch.MessageID:
                    Append("body.peek[header.fields (Message-ID)]");
                    break;
                default:
                    Append("body.peek[]");
                    break;
            }
            Append("\r\n");
        }

        /*
            * 1 FETCH (BODY[]<0> {10}
        */
        protected override async Task<Return> ParseInternal(InputType t, string str)
        {
            if (t == InputType.Line)
            {
                Regex re = new Regex(@"^\* ([0-9]+) fetch [^{]+[{]([0-9]+)[}]$", RegexOptions.IgnoreCase);
                Match m = re.Match(str);
                if (m.Success)
                {
                    msgnum = m.Groups[1].Value;
                    return (new Return(MoreInput.Chunk, int.Parse(m.Groups[2].Value)));
                }
                else
                    return (new Return(MoreInput.Line));
            }
            else
            {
                await cb(str,msgnum);
                return (new Return(MoreInput.Line));
            }
        }
    }

    class NullCommand : Command
    {
        public NullCommand() : base(null, true)
        {
        }

        protected override async Task<Return> ParseInternal(InputType t, string str)
        {
            await Task.Yield();
            return (new Return(MoreInput.Done));
        }
    }

    class ImapCmd
    {
        StreamReader reader { get; set; }
        Stream writer { get; set; }
        Command command { get; set; }
        bool compress { get; set; }
        CancellationToken token { get; set; }

        public ImapCmd(CancellationToken token, StreamReader reader, Stream writer, bool compress = false)
        {
            this.reader = reader;
            this.writer = writer;
            this.compress = compress;
            this.token = token;
        }

        async Task Write(String str)
        {
            byte[] buffer = Encoding.Default.GetBytes(str);
            if (compress)
            {
                MemoryStream mems = new MemoryStream();
                DeflateStream defl = new DeflateStream(mems, CompressionLevel.Fastest, true);
                defl.Write(buffer, 0, buffer.Length);
                defl.Close();
                buffer = mems.GetBuffer();
            }
            await writer.WriteAsync(buffer, 0, buffer.Length);
            await writer.FlushAsync();
        }

        public async Task<T> WithTimeout<T>(Task<T> task, int msec)
        {
            if (await Task.WhenAny(task, Task.Delay(msec, token)) == task)
            {
                if (task.IsCompleted)
                    return task.Result;
                else
                    throw new ServerIOException("Server closed connection");
            }
            else if (token.IsCancellationRequested)
                throw new CancelException();
            else
                throw new TimeoutException("Network read timeout");
        }

        public async Task<ReturnCode> Run(Command command)
        {
            int timeout = 5 * 60 * 1000;//msec(5 min)
            try
            {
                if (command.Cmd != "")
                   await Write(command.Cmd);
                // read until Ok,Bad,No, if there is literal at the end then read the block
                for (Command.Return ret = new Command.Return(Command.MoreInput.Line); ;)
                {
                    if (token.IsCancellationRequested)
                    {
                        command.Ret = ReturnCode.Cancelled;
                        return command.Ret;
                    }
                    switch (ret.Code)
                    {
                        case Command.MoreInput.Line:
                            String str = await WithTimeout(reader.ReadLineAsync(), timeout);
                            if (str == null)
                                throw new ServerIOException("Server closed connection");
                            ret = await command.Parse(Command.InputType.Line, str);
                            break;
                        case Command.MoreInput.Chunk:
                            int count = ret.Size.GetValueOrDefault();
                            char[] buff = new char[count];
                            for (int offset = 0, read = 0; count != 0; offset += read, count -= read)
                            {
                                read = await WithTimeout(reader.ReadBlockAsync(buff, offset, count), timeout);
                                if (read == 0)
                                    throw new ServerIOException("Server closed connection");
                            }
                            ret = await command.Parse(Command.InputType.Chunk, new String(buff));
                            break;
                        default:
                            return command.Ret;
                    }
                }
            }
            finally
            {
            }
        }
    }
}
