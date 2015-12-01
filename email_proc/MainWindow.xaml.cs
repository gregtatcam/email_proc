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
using System.Windows;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;


namespace email_proc
{
    class ValidationException : Exception
    {
        public String reason { get; set; }
        public ValidationException(String reason)
        {
            this.reason = reason;
        }
    }
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<String,String> imapAddr = new Dictionary<string, string>();
        Key lastDebug = Key.System;
        DateTime now = DateTime.Now;
        public MainWindow()
        {
            InitializeComponent();
            cbAddr.Focus();
            String[] imap = new String [] { "163","aol", "att", "comcast","cox","gmail","gmx","hermes", "icloud","inbox","mail","optimum","outlook",
                "rambler", "yahoo", "yandex", "yeah", "zoho"};
            imapAddr.Add("163", "imap.163.com");
            imapAddr.Add("aol", "imap.aol.com");
            imapAddr.Add("att", "imap.mail.att.com");
            imapAddr.Add("comcast", "imap.comcast.net");
            imapAddr.Add("cox", "imap.cox.net");
            imapAddr.Add("gmail", "imap.gmail.com");
            imapAddr.Add("gmx", "imap.gmx.com");
            imapAddr.Add("hermes", "imap.hermes.cam.ac.uk");
            imapAddr.Add("icloud", "imap.mail.me.com");
            imapAddr.Add("inbox", "imap.inbox.com");
            imapAddr.Add("mail", "imap.mail.com");
            imapAddr.Add("optimum", "mail.optimum.net");
            imapAddr.Add("outlook", "imap-mail.outlook.com");
            imapAddr.Add("rambler", "imap.rambler.ru");
            imapAddr.Add("yahoo", "imap.mail.me.com");
            imapAddr.Add("yandex", "imap.yandex.com");
            imapAddr.Add("yeah", "imap.yeah.net");
            imapAddr.Add("zoho", "imap.zoho.com");

            cbAddr.ItemsSource = imap;
        }

        void Validate()
        {
            String error = "Email archive must be entered";
            if ((bool)cbDownload.IsChecked)
            {
                if (cbAddr.ToString() == "")
                    throw new ValidationException("Server address must be entered");
                if (txtUser.Text == "")
                    throw new ValidationException("User must be entered");
                if (txtPassword.Password == "")
                    throw new ValidationException("Password must be entered");
                error = "Download directory must be entered";
            }
            if (txtDownload.Text == "")
                throw new ValidationException(error);
        }

        private async void btnStart_ClickTest(object sender, RoutedEventArgs e)
        {
            //MessageReader reader = new MessageReader(@"c:\Downloads\arch130927535041968465.mbox");
            MessageReader reader = new MessageReader(new MemoryStream(Encoding.ASCII.GetBytes(Tests.NestedMultipart())));
            await EmailParser.ParseMessages(reader, async delegate (Message message)
            {
                String postmark = message.postmark.GetString();
                await Task.Run(() => { EmailParser.TraverseEmail(message.email); });
            });
            return;
        }

        void Status(bool cr, String format, params object[] args)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(format, args);
            if (cr)
                lbStatus.Items[lbStatus.Items.Count - 1] = sb.ToString();
            else
                lbStatus.Items.Add(sb.ToString());
            lbStatus.Items.Refresh();
            lbStatus.ScrollIntoView(lbStatus.Items[lbStatus.Items.Count - 1]);
        }

        async Task SaveMessage (StreamWriter filew, String mailbox, String message, String msgnum)
        {
            try
            {
                String status = (String)lbStatus.Items[lbStatus.Items.Count - 1];
                Regex re = new Regex("^((\"[^\"]+\")|([^\" ]+) - [0-9]+ messages: )");
                Match m = re.Match(status);
                if (m.Success)
                {
                    Status(true, "{0} {1}", m.Groups[1].Value, msgnum);
                }
                String postmark = EmailParser.MakePostmark(message);

                await filew.WriteAsync(postmark);
                await filew.WriteAsync("\r\n");
                await filew.WriteAsync("X-Mailbox: " + mailbox);
                await filew.WriteAsync("\r\n");
            }
            catch (AlreadyExistsException)
            {
            }
            finally
            {
                await filew.WriteAsync(message);
            }
        }

        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            StreamWriter filew = null;
            String file = null;
            String dir = null;
            try {
                Validate();
                lbStatus.Items.Clear();
                lbStatus.Items.Refresh();
                btnStart.IsEnabled = false;
                btnBrowse.IsEnabled = false;
                StringBuilder sb = new StringBuilder();
                if ((bool)cbDownload.IsChecked)
                {
                    dir = txtDownload.Text;
                    sb.AppendFormat(@"{0}\arch{1}.mbox", dir, DateTime.Now.ToFileTime());
                    file = sb.ToString();
                    sb.Clear();
                    sb.AppendFormat(@"{0}\email_proc1596.index", dir);
                    bool resume = false;
                    if (File.Exists(sb.ToString()))
                    {
                        resume = true;
                        file = File.ReadAllText(sb.ToString());
                    }
                    else
                        File.WriteAllText(sb.ToString(), file);
                    String addr = cbAddr.Text;
                    if (imapAddr.ContainsKey(addr))
                        addr = imapAddr[addr];
                    StateMachine sm = new StateMachine(addr, int.Parse(txtPort.Text), txtUser.Text, txtPassword.Password, file);
                    DateTime start_time = DateTime.Now;
                    await sm.Start(
                        delegate (String format, object[] args)
                        {
                            Status(false, format, args);
                        },
                        async delegate (String mailbox, String message, String msgnum) 
                            {
                                if (filew == null)
                                    filew = new StreamWriter(file, resume);
                                await SaveMessage(filew, mailbox, message, msgnum);
                            },
                        delegate (double progress) { prBar.Value = progress; }
                    );
                    Status(false, "Downloaded email to {0}", file);
                    TimeSpan span = DateTime.Now - start_time;
                    Status(false, "Download time {0} seconds", span.TotalSeconds);
                    filew.Close();
                }
                else
                {
                    file = txtDownload.Text;
                    dir = Path.GetDirectoryName(txtDownload.Text);
                }
                if ((bool)cbStatistics.IsChecked)
                {
                    EmailStats stats = new EmailStats();
                    await stats.Start(dir, file, Status, delegate(double progress) { prBar.Value = progress; });
                    prBar.Value = 100;
                }
            }
            catch (FailedLoginException)
            {
                Status(false, "Login failed, invalid user or password");
            }
            catch (FailedException)
            {
                Status(false, "Failed to download");
            }
            catch (SslFailedException)
            {
                Status(false, "Failed to establish secure connection");
            }
            catch (ValidationException ex)
            {
                Status(false, ex.reason);
            }
            catch (Exception ex)
            {
                Status(false, ex.Message);
            }
            finally
            {
                prBar.IsIndeterminate = false;
                btnBrowse.IsEnabled = true;
                btnStart.IsEnabled = true;
                if (filew != null)
                    filew.Close();
            }
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)cbDownload.IsChecked == true)
            {
                FolderBrowserDialog dlg = new FolderBrowserDialog();
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    txtDownload.Text = dlg.SelectedPath;
            } else
            {
                OpenFileDialog dlg = new OpenFileDialog();
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    txtDownload.Text = dlg.FileName;
            }
        }

        void SetControls (bool download)
        {
            txtPassword.IsEnabled = download;
            txtPort.IsEnabled = download;
            txtUser.IsEnabled = download;
            cbAddr.IsEnabled = download;
            if (download)
                lblDownload.Content = "Download directory:";
            else
                lblDownload.Content = "Email archive:";
        }

        private void cbDownload_Checked(object sender, RoutedEventArgs e)
        {
            SetControls(true);
        }

        private void cbStatistics_Checked(object sender, RoutedEventArgs e)
        {
        }

        private void cbDownload_Unchecked(object sender, RoutedEventArgs e)
        {
            SetControls(false);
            cbStatistics.IsChecked = true;
        }

        private void cbStatistics_Unchecked(object sender, RoutedEventArgs e)
        {
            cbDownload.IsChecked = true;
            SetControls(true);
        }

        private void txtDownload_GotFocus(object sender, RoutedEventArgs e)
        {
            txtDownload.ToolTip = txtDownload.Text;
        }

        private void Grid_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.D && (DateTime.Now - now).Seconds < 2 && lastDebug == Key.LeftCtrl)
            {
                cbDownload.Visibility = Visibility.Visible;
                cbStatistics.Visibility = Visibility.Visible;
            }
            lastDebug = e.Key;
            now = DateTime.Now;
        }
    }
}
