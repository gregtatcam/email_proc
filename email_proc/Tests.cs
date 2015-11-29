using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace email_proc
{
    class Tests
    {
        public static String MultipartPreamble()
        {
            return @"From: Nathaniel Borenstein <nsb@bellcore.com>
To: Ned Freed <ned@innosoft.com>
Date: Sun, 21 Mar 1993 23:56:48 -0800 (PST)
Subject: Sample message
MIME-Version: 1.0
Content-type: multipart/mixed; boundary=""simple boundary""

This is the preamble.It is to be ignored, though it
is a handy place for composition agents to include an
explanatory note to non - MIME conformant readers.

--simple boundary

This is implicitly typed plain US - ASCII text.
It does NOT end with a linebreak.
--simple boundary
Content-type: text/plain; charset=us-ascii

This is explicitly typed plain US - ASCII text.
It DOES end with a linebreak.

--simple boundary--

This is the epilogue.It is also to be ignored.";

        }

        public static String NestedMultipart()
        {
            return
@"From 1487928187900928398@xxx Fri Dec 19 02:21:37 2014\r
Content-Type: multipart/mixed; boundary=42

--42

plain text 1
--42

plain text 2
--42
Content-Type: multipart/mixed; boundary=43

--43

plain text 1.1
--43

plain text 1.2
--43--

--42
Content-Type: message/rfc822

Subject: test
From: john.dow@cloud.net
Content-Type: multipart/mixed; boundary=44

--44

Some text in rfc822
--44
Content-Type: vidieo/jpg

abcdefg
hijklmn
opqrstu
wxyz

--44--

--42--";
        }
    }
}
