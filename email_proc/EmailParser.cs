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
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;

namespace email_proc
{
    public enum DataType { Data, Message, Multipart }
    public enum ContentType { Audio, Video, Image, Application, Text, Multipart, Message, Other }
    public enum ContentSubtype { Plain, Rfc822, Digest, Alternative, Parallel, Mixed, Other }
    public enum ParseResult { Ok, OkMultipart, Eof, Failed }

    // so far used to cache the postmark if the reader consumed too much data
    public class MessageReader : StreamReader
    {
        Queue<String> cache { get; set; }
        public MessageReader(String path) : base(path)
        {
            cache = new Queue<string>();
        }
        public MessageReader(Stream strm) : base(strm)
        {
            cache = new Queue<String>();
        }
        public MessageReader(Stream strm, Encoding encoding, bool detectEncoding, int bufferSize, bool leaveOpen) :
            base(strm, encoding, detectEncoding, bufferSize, leaveOpen)
        {
            cache = new Queue<String>();
        }
        public new Task<String> ReadLineAsync()
        {
            if (cache.Count != 0)
            {
                return Task.FromResult<String>(cache.Dequeue());
            }
            else
                return base.ReadLineAsync();
        }
        public void PushCacheLine(String str)
        {
            cache.Enqueue(str);
        }
    }

    class ParsingFailedException : Exception
    {
        public ParsingFailedException(String message) : base(message)
        {
        }
    }
    class AlreadyExistsException : Exception
    {
        public AlreadyExistsException() { }
    }

    public class Boundary
    {
        static Regex re_boundary = new Regex("boundary=((\"[^\"]+\")|([^ ]+))", RegexOptions.IgnoreCase);
        String boundary { get; set; }
        public String openBoundary { get; private set; }
        public String closeBoundary { get; private set; }
        public Boundary(String delimeter=null)
        {
            if (delimeter != null)
            {
                delimeter = delimeter.Trim('\"').TrimEnd(' ');
                openBoundary = "--" + delimeter;
                closeBoundary = openBoundary + "--";
            }
            else
                boundary = null;
        }
        public bool IsOpen(String line)
        {
            return line.TrimEnd(' ') == openBoundary;
        }
        public bool IsClose(String line)
        {
            return line.TrimEnd(' ') == closeBoundary;
        }
        public bool IsEmpty()
        {
            return boundary == null;
        }
        public static Boundary Parse(String line)
        {
            Match m = re_boundary.Match(line);
            if (m.Success)
                return new Boundary(m.Groups[1].Value);
            return null;
        }
    }

    public abstract class Entity
    {
        static String crlf = "\n";
        protected MemoryStream entity { get; set; }
        protected long position { get; set; }
        public int size { get; protected set; }
        public int lines { get; protected set; }
        // assume the entity starts at the current position of the stream
        public Entity(MemoryStream entity)
        {
            if (entity == null)
                entity = new MemoryStream();
            position = entity.Position;
            size = 0;
            this.entity = entity;
        }

        private void Write(String buffer)
        {
            using (Lock())
            {
                entity.WriteAsync(Encoding.ASCII.GetBytes(buffer), 0, buffer.Length);
            }
        }

        protected void WriteWithCrlf(String buffer)
        {
            StringBuilder sb = new StringBuilder(buffer);
            sb.Append(crlf);
            Write(sb.ToString());
            lines++;
        }

        protected void WriteCrlf()
        {
            Write(crlf);
            lines++;
        }
        protected Mutex Lock()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(entity.GetHashCode());
            bool created = false;
            Mutex nmutex = new Mutex(true, sb.ToString(), out created);
            if (created == false)
                nmutex.WaitOne();
            return nmutex;
        } 
        public String GetString()
        {
            return Encoding.ASCII.GetString(GetBytes());
        }
        public byte[] GetBytes ()
        {
            using (Lock())
            {
                entity.Position = position;
                byte[] buffer = new byte[size];
                for (int index = 0, count = size, read = 0; count != 0; index += read, count -= read)
                    read = entity.Read(buffer, index, count);
                return buffer;
            }
        }
        public void SetSize()
        {
            using (Lock())
            {
                size = (int)(entity.Position - position);
            }
        }
        public void RewindLastCrlfSize()
        {
            using (Lock())
            {
                entity.Seek(-2, SeekOrigin.Current);
                byte[] buff = new byte[2];
                entity.Read(buff, 0, 2);
                char[] crlf = Encoding.ASCII.GetChars(buff);
                if (crlf[0] == '\r')
                    size--;
                if (crlf[1] == '\n')
                    size--;
            }
        }

        protected async Task ConsumeToEnd(MessageReader reader)
        {
            ParseResult res = ParseResult.Ok;
            while (res != ParseResult.Eof)
            {
                String line = await reader.ReadLineAsync();
                if (line == null)
                    break;
                else if (EmailParser.IsPostmark(line))
                {
                    reader.PushCacheLine(line);
                    break;
                }
                else
                    WriteWithCrlf(line);
            }
        }
    }

    public class Headers : Entity
    {
        public delegate bool HeaderCb(String name, String value);
        public ContentType contentType { get; private set; }
        public ContentSubtype contentSubtype { get; private set; }
        public String contentTypeFullStr { get; private set; }
        public Boundary boundary { get; private set; }
        static Dictionary<string, ContentType> types = new Dictionary<string, ContentType>()
            { { "audo", ContentType.Audio }, {"video", ContentType.Video}, {"image",ContentType.Image }, {"application",ContentType.Application },
            {"multipart",ContentType.Multipart }, {"message",ContentType.Message } };
        static Dictionary<string, ContentSubtype> subtypes = new Dictionary<string, ContentSubtype>()
            { { "plain", ContentSubtype.Plain }, {"rfc822",ContentSubtype.Rfc822 }, {"digest",ContentSubtype.Digest },
              { "alternative", ContentSubtype.Alternative }, { "parallel", ContentSubtype.Parallel }, { "mixed", ContentSubtype.Mixed} };
        static Regex re_content = new Regex("^content-type: ([^/ ]+)/([^; ]+)(.*)$", RegexOptions.IgnoreCase);
        public Headers(MemoryStream entity, ContentType outerType = ContentType.Text, ContentSubtype outerSubtype = ContentSubtype.Plain):base(entity)
        {
            contentTypeFullStr = "";
            if (outerType == ContentType.Multipart && outerSubtype == ContentSubtype.Digest)
            {
                contentType = ContentType.Message;
                contentSubtype = ContentSubtype.Rfc822;
            } else
            {
                contentType = ContentType.Text;
                contentSubtype = ContentSubtype.Plain;
            }
            boundary = null;
        }

        public void GetHeaders(HeaderCb cb)
        {
            String name = "";
            String value = "";
            bool done = false;

            using (StringReader reader = new StringReader(GetString()))
            {
                String line = "";
                while (done == false && (line = reader.ReadLine()) != null)
                {
                    Match m = Regex.Match(line, "^([^\t :]+):(.*)$");
                    // matched header: field
                    if (m.Success)
                    {
                        // call with last consumed header/value. new header means all (if any) values with FWS are consumed
                        if (name != "")
                            done = cb(name, value);
          
                        name = m.Groups[1].Value.ToLower();
                        value = m.Groups[2].Value.TrimStart(' ');
                    }
                    // must be FWS
                    else if (name != "")
                    {
                        value += " " + line.TrimStart(' ');
                    }
                }
            }

            if (name != "" && done == false)
                cb(name, value);
        }
        
        // if hedadersToGet is null then get all headers
        public Dictionary<String,String> GetDictionary(Dictionary<String,String> headersToGet=null)
        {
            Dictionary<String, String> headers = new Dictionary<string, string>();
            GetHeaders(delegate (String name, String value)
            {
                if (headersToGet == null || headersToGet.ContainsKey(name))
                    headers[name] = value.Trim();
                else if (headersToGet != null && headersToGet.Count == headers.Count) // must have got all requested
                    return true;
                return false;
            });
            foreach (String name in headersToGet.Keys)
            {
                if (headers.ContainsKey(name) == false)
                    headers[name] = "";
            }
            return headers;
        }

        public int GetNumberOfHeaders()
        {
            int n = 0;
            GetHeaders(delegate (String nm, String vl) { n++; return false; });
            return n;
        }

        public async Task<ParseResult> Parse(MessageReader reader)
        {
            // parse until empty line, which start the body
            String line = "";
            bool foundContentType = false;
            bool boundaryRequired = false;
            // could there be an empty line in FWS? I think just crlf is not allowed in FWS
            // if starts with the blank line then there is no header
            // ends with blank line
            while ((line = await reader.ReadLineAsync()) != null && line != "")
            {
                WriteWithCrlf(line);
                if (foundContentType && boundaryRequired && boundary == null)
                {
                    boundary = Boundary.Parse(line);
                }
                else if (foundContentType == false)
                {
                    Match m = re_content.Match(line);
                    if (m.Success)
                    {
                        ContentType type = ContentType.Text;
                        ContentSubtype subtype = ContentSubtype.Plain;
                        String tp = m.Groups[1].Value.ToLower();
                        String sbtp = m.Groups[2].Value.ToLower();
                        contentTypeFullStr = tp + "/" + sbtp;
                        if (types.TryGetValue(tp, out type) == true)
                            contentType = type;
                        if (subtypes.TryGetValue(sbtp, out subtype) == true)
                            contentSubtype = subtype;
                        foundContentType = true;
                        if (contentType == ContentType.Multipart)
                        {
                            boundaryRequired = true;
                            boundary = Boundary.Parse(m.Groups[3].Value);
                        }
                    }
                }
            }
            if (boundaryRequired && boundary == null)
                throw new ParsingFailedException("multipart media part with no boundary");
            SetSize();
            WriteCrlf(); // delimeter between headers and body, not part of the headers, so not included in size
            if (line == null)
                return ParseResult.Eof;
            else
                return ParseResult.Ok;
        }
    }

    public class Content : Entity
    {
        public DataType dataType { get; private set; }
        public List<Email> data { get; private set; }
        public Content(MemoryStream entity) : base(entity)
        {
            data = null;
        }
        void Add(Email email)
        {
            if (data == null)
                data = new List<Email>();
            data.Add(email);
        }

        public async Task<ParseResult> Parse(MessageReader reader, ContentType type = ContentType.Text,
            ContentSubtype subtype = ContentSubtype.Plain, Boundary boundary = null)
        {
            if (type == ContentType.Multipart)
            {
                dataType = DataType.Multipart;
                while (true)
                {
                    String line = await reader.ReadLineAsync();
                    if (line == null)
                    {
                        SetSize();
                        return ParseResult.Eof;
                    }
                    else if (EmailParser.IsPostmark(line))
                    {
                        // consumed too much, probably missing boundary?
                        reader.PushCacheLine(line);
                        SetSize();
                        return ParseResult.Eof;
                    }
                    WriteWithCrlf(line);
                    // find open boundary
                    if (boundary.IsOpen(line))
                    {
                        Email email = null;
                        ParseResult res;
                        do
                        {
                            // consume all parts, consisting of header (optional) and content
                            // the boundary token delimets the part
                            // the close boundary completes multipart parsing
                            // content in the multipart is responsible for consuming it's delimeter (end)
                            // exception is the last part which is also multipart
                            email = new Email(entity);
                            Add(email);
                        } while ((res = await email.Parse(reader, type, subtype, boundary)) == ParseResult.Ok);
                        // if the last part is a multipart itself then it doesn't consume the close boundary
                        // or more parts, continue parsing until all parts and close boundary are consumed
                        if (data.Last<Email>().content.dataType == DataType.Multipart)
                            continue;
                        SetSize();
                        return res;
                    }
                    else if (boundary.IsClose(line))
                    {
                        SetSize();
                        return ParseResult.OkMultipart;
                    }
                }
            }
            else if (type == ContentType.Message)
            {
                dataType = DataType.Message;
                Email email = new Email(entity);
                Add(email);
                ParseResult res = await email.Parse(reader, type, subtype, boundary);
                SetSize();
                return res;
            }
            else
            {
                dataType = DataType.Data;
                while (true)
                {
                    String line = await reader.ReadLineAsync();
                    if (line == null)
                    {
                        SetSize();
                        return ParseResult.Eof;
                    }
                    else if (EmailParser.IsPostmark(line))
                    {
                        // consumed too much, probably closing boundary is missing ?
                        reader.PushCacheLine(line);
                        SetSize();
                        return ParseResult.Ok;
                    }
                    else if (boundary != null && boundary.IsOpen(line))
                    {
                        SetSize();
                        RewindLastCrlfSize();
                        WriteWithCrlf(line);
                        return ParseResult.Ok;
                    }
                    else if (boundary != null && boundary.IsClose(line))
                    {
                        SetSize();
                        RewindLastCrlfSize();
                        WriteWithCrlf(line);
                        return ParseResult.OkMultipart;
                    }
                    else
                        WriteWithCrlf(line);
                }
            }
        }
    }

    public class Email : Entity
    {
        public Headers headers { get; private set; }
        public Content content { get; private set; }
        public Email(MemoryStream entity) : base(entity)
        {
        }
        public async Task<ParseResult> Parse(MessageReader reader, ContentType type = ContentType.Text,
            ContentSubtype subtype = ContentSubtype.Plain, Boundary boundary = null)
        {
            headers = new Headers(entity, type, subtype);
            if ((await headers.Parse(reader)) == ParseResult.Failed)
                throw new ParsingFailedException("email doesn't contain headers");
            content = new Content(entity);
            ParseResult result = await content.Parse(reader, headers.contentType, headers.contentSubtype,
                (headers.boundary != null) ? headers.boundary : boundary);
            if (result == ParseResult.Failed)
                throw new ParsingFailedException("failed to parse email body");
            return result;
        }
    }

    public class Postmark: Entity
    {
        public Postmark (MemoryStream entity) : base(entity)
        {
        }

        public async Task<ParseResult> Parse(MessageReader reader)
        {
            String line = null;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                if (EmailParser.IsPostmark(line))
                {
                    WriteWithCrlf(line);
                    // assume the Postmark always starts with 0 position
                    size = (int)entity.Position;
                    return ParseResult.Ok;
                }
            }
            return ParseResult.Failed;
        }
    }

    public class Message: Entity
    {
        public Postmark postmark { get; private set; }
        public Email email { get; private set; }

        public Message() : base(null)
        {
        }

        /* Assume (for now) the message starts with the postmark,
           Further assume the message structure
           postmark\r\n
           headers\r\n
           \r\n
           body
        */
        public async Task<ParseResult> Parse(MessageReader reader)
        {
            postmark = new Postmark(entity);
            if ((await postmark.Parse(reader)) == ParseResult.Failed)
                throw new ParsingFailedException("postmark is not found");
            email = new Email(entity);
            ParseResult res = await email.Parse(reader);
            if (res == ParseResult.Failed)
                throw new ParsingFailedException("email doesn't conform to rfc822");
            if (res != ParseResult.Eof)
                await ConsumeToEnd(reader);
            SetSize();

            return res;
        }
    }

    public class EmailParser
    {
        public delegate void TraverseCb(Email email);
        public delegate Task MessageCb(Message message, Exception ex=null);
        // From 1487928187900928398@xxx Fri Dec 19 02:21:37 2014
        public static String PostmarkReStr = "^(from [^ \r\n]+ (mon|tue|wed|thu|fri|sat|sun) (jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec))[^\r\n]+";
        // From: John Smith <john.smith@provider.com>
        public static String FromReStr = "^from: ((\"[^\"\r\n]+\"[\r\n]*)|(([^:\"@<]+[\r\n]*){0,}))[ ]*<?([^\"<: \r\n@]+@[^\": \r\n>]+)>?([\r\n]*)";
        // Date: Fri, 19 Dec 2014 09:21:23 -0500
        public static String DateReStr = "^date: (mon|tue|wed|thu|fri|sat|sun), ([0-9]+) (jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec) ([0-9]+) ([0-9:]+)";
        public EmailParser()
        {
        }

        public static bool IsPostmark(String line)
        {
            return Regex.IsMatch(line, PostmarkReStr, RegexOptions.IgnoreCase);
        }

        public static String MakePostmark(String message)
        {
            if (Regex.IsMatch(message, PostmarkReStr, RegexOptions.IgnoreCase))
                throw new AlreadyExistsException();

            message = message.Substring(0, message.Length > 5000 ? 5000 : message.Length);
                  
            Match m = Regex.Match(message, FromReStr, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            StringBuilder sb = new StringBuilder("From ");
            String from = "daemon@local.com";
            if (m.Success)
                from = m.Groups[5].Value;
            sb.Append(from);
            String date = DateTime.Today.ToString("ddd MMM dd HH:mm:ss yyyy");
            m = Regex.Match(message, DateReStr, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            if (m.Success)
            {
                String dow = m.Groups[1].Value;
                String day = m.Groups[2].Value;
                String mon = m.Groups[3].Value;
                String year = m.Groups[4].Value;
                String time = m.Groups[5].Value;
                StringBuilder sb1 = new StringBuilder();
                sb1.AppendFormat("{0} {1} {2} {3} {4}", dow, mon, day, time, year);
                date = sb1.ToString();
            }
            sb.AppendFormat(" {0}", date);
            return sb.ToString();
        }

        public async static Task ParseMessages(CancellationToken token, MessageReader reader, MessageCb cb)
        {
            while (reader.EndOfStream == false)
            {
                if (token.IsCancellationRequested)
                    break;
                try
                {
                    Message message = new Message();
                    if (await message.Parse(reader) != ParseResult.Failed)
                        await cb(message);
                }
                catch (Exception ex)
                {
                    await cb(null, ex);
                }
            }
            return;
        }

        public async static Task ParseToResume(CancellationToken token, String addr, String user, String file, StatusCb status, ProgressCb progress)
        {
            String index = Path.Combine(Path.GetDirectoryName(file), "email_proc1596.index");
            FileInfo info = new FileInfo(file);
            long length = info.Length;
            FileStream stream = info.OpenRead();
            status("Generating resume file...");
            using (StreamReader reader = new StreamReader(stream))
            {
                using (StreamWriter writer = new StreamWriter(index))
                {
                    await writer.WriteLineAsync(addr + " " + user + " " + file);
                    String line = "";
                    String mailbox = "";
                    String messageid = "";
                    Regex re_mailbox = new Regex("^X-Mailbox: (.+)$");
                    Regex re_messageid = new Regex("^Message-ID: ([^ ]+)", RegexOptions.IgnoreCase);
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        if (token.IsCancellationRequested)
                            return;
                        progress(100.0 * stream.Position / length);
                        Match m = re_mailbox.Match(line);
                        if (m.Success)
                        {
                            if (mailbox != "")
                            {
                                if (messageid != "")
                                    messageid = " " + messageid;
                                await writer.WriteLineAsync(mailbox + messageid);
                            }
                            mailbox = m.Groups[1].Value;
                            messageid = "";
                        }
                        else if (messageid == "")
                        {
                            m = re_messageid.Match(line);
                            if (m.Success)
                                messageid = m.Groups[1].Value;
                        }
                    }
                    if (mailbox != "")
                    {
                        if (messageid != "")
                            messageid = " " + messageid;
                        await writer.WriteLineAsync(mailbox + messageid);
                    }
                }
            }
        }

        public static void TraverseEmail(Email email)
        {
            String headers = email.headers.GetString();
            ContentType ctype = email.headers.contentType;
            DataType dtype = email.content.dataType;
            int size = email.content.size;
            switch (dtype)
            {
                case DataType.Data:
                    String content = email.content.GetString();
                    break;
                case DataType.Message:
                case DataType.Multipart:
                    foreach (Email e in email.content.data)
                        TraverseEmail(e);
                    break;
                default:
                    break;
            }
            return;
        }
    }
}
