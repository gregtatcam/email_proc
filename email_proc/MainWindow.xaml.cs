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
using System.Threading;
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
        CancellationTokenSource cancelSrc = null;
        int tabCtrlSelected = 0;
        DateTime start_time;
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
            imapAddr.Add("yahoo", "imap.mail.yahoo.com");
            imapAddr.Add("yandex", "imap.yandex.com");
            imapAddr.Add("yeah", "imap.yeah.net");
            imapAddr.Add("zoho", "imap.zoho.com");

            txtDownload.Text = Path.Combine("C:\\Users", Environment.UserName, "Downloads");

            cbAddr.ItemsSource = imap;
            cbStatistics.Visibility = Visibility.Hidden;
            cbDownload.Visibility = Visibility.Hidden;
            cbResume.Visibility = Visibility.Hidden;

            OtherControls(Visibility.Hidden);
        }

        void OtherControls(Visibility other)
        {
            lblOtherImap.Visibility = other;
            lblOtherPort.Visibility = other;
            txtOtherImap.Visibility = other;
            txtOtherPort.Visibility = other;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            cancelSrc.Cancel();
            Thread.Sleep(10000);
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

        void Status(bool cr, string crcondition, String format, params object[] args)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(format, args);
            if (cr && ("" == crcondition || Regex.IsMatch(lbStatus.Items[lbStatus.Items.Count - 1].ToString(), crcondition)))
                lbStatus.Items[lbStatus.Items.Count - 1] = sb.ToString();
            else
                lbStatus.Items.Add(sb.ToString());
            lbStatus.Items.Refresh();
            lbStatus.ScrollIntoView(lbStatus.Items[lbStatus.Items.Count - 1]);
        }

        async Task SaveMessage (StreamWriter filew, StreamWriter indexw, String mailbox, String message, String msgid, String msgnum)
        {
            try
            {
                String status = (String)lbStatus.Items[lbStatus.Items.Count - 1];
                Regex re = new Regex("^((\"[^\"]+\")|([^\" ]+) - [0-9]+ messages: )");
                Match m = re.Match(status);
                if (m.Success)
                {
                    Status(true, "", "{0} {1}", m.Groups[1].Value, msgnum);
                }
                String postmark = EmailParser.MakePostmark(message);

                await filew.WriteAsync(postmark);
                await filew.WriteAsync("\r\n");
                await filew.WriteAsync("X-Gmail-Labels: " + mailbox);
                await filew.WriteAsync("\r\n");
                await indexw.WriteLineAsync(mailbox + (msgid != "" ? " " + msgid : ""));
            }
            catch (AlreadyExistsException)
            {
            }
            finally
            {
                await filew.WriteAsync(message);
            }
        }

        async Task<bool> Resume(String addr)
        {
            if (File.Exists(txtDownload.Text) == false)
            {
                System.Windows.Forms.MessageBox.Show("Enter file to resume download");
                return false;
            }
            await EmailParser.ParseToResume(cancelSrc.Token, addr, txtUser.Text, txtDownload.Text, (f,o) => { Status(false, "", f, o); }, (v) => { prBar.Value = v; });
            return (cancelSrc.Token.IsCancellationRequested == false);
        }

        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if ((String)btnStart.Content == "Cancel")
            {
                cancelSrc.Cancel();
                btnStart.Content = "Start";
                btnBrowse.IsEnabled = true;
                return;
            }
            start_time = DateTime.Now;
            await Start(1);
        }

        private async Task Start(int attempts)
        {
            if (attempts > 200)
            {
                Status(false, "", "Stopped, too many retry attempts. Click on Start to resume the download");
                return;
            }
            else if (attempts > 1)
                Status(false, "", "Retrying...");
            cancelSrc = new CancellationTokenSource();
            StreamWriter filew = null;
            StreamWriter indexw = null;
            String file = null;
            String dir = null;
            try {
                Validate();
                if (attempts == 1)
                {
                    lbStatus.Items.Clear();
                    lbStatus.Items.Refresh();
                }
                btnStart.Content = "Cancel";
                btnBrowse.IsEnabled = false;
                StringBuilder sb = new StringBuilder();
                bool dldNotDone = false;
                if ((bool)cbDownload.IsChecked)
                {
                    String addr = cbAddr.Text;
                    if (imapAddr.ContainsKey(addr))
                        addr = imapAddr[addr];

                    dir = txtDownload.Text;

                    if (cbResume.IsChecked == true)
                    {
                        if (await Resume(addr) == false)
                            return;
                        dir = Path.GetDirectoryName(dir);
                    }

                    sb.AppendFormat(@"{0}\arch{1}.mbox", dir, DateTime.Now.ToFileTime());
                    file = sb.ToString();
                    String indexFile = Path.Combine(dir, "email_proc1596.index");
                    bool resume = false;
                    String downloadedFile = "";
                    if (File.Exists(indexFile))
                    {
                        using (StreamReader reader = new StreamReader(indexFile))
                        {
                            String line = reader.ReadLine();
                            if (line != null)
                            {
                                Match m = Regex.Match(line, "^([^ ]+) ([^ ]+) (.+)$");
                                if (m.Success && m.Groups[1].Value == addr && m.Groups[2].Value == txtUser.Text)
                                {
                                    downloadedFile = m.Groups[3].Value;
                                    file = downloadedFile;
                                }
                            }
                        }
                    }
                    if (downloadedFile == "")
                    {
                        using (StreamWriter writer = new StreamWriter(indexFile))
                        {
                            writer.WriteLine(addr + " " + txtUser.Text + " " + file);
                        }
                    }
                    else
                    {
                        file = downloadedFile;
                        resume = true;
                    }
                    filew = new StreamWriter(file, resume);
                    StateMachine sm = new StateMachine(cancelSrc.Token, addr, int.Parse(txtPort.Text), txtUser.Text, txtPassword.Password, file);
                    try
                    {
                        await sm.Start(
                            delegate (String format, object[] args)
                            {
                                Status(false, "", format, args);
                            },
                            async delegate (String mailbox, String message, String msgid, String msgnum)
                            {
                                if (indexw == null)
                                    indexw = new StreamWriter(indexFile, true);
                                await SaveMessage(filew, indexw, mailbox, message, msgid, msgnum);
                            },
                            delegate (double progress) { prBar.Value = progress; }
                        );
                    }
                    catch (ServerIOException ex)
                    {
                        Status(false, "", ex.Message);
                    }
                    catch (TimeoutException ex)
                    {
                        Status(false, "", ex.Message);
                    }
                    catch (FailedException)
                    {
                        Status(false, "", "Failed to download");
                    }
                    catch (ConnectionFailedException ex)
                    {
                        Status(false, "", ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Status(false, "", ex.Message);
                    }
                    finally
                    {
                        if (indexw != null)
                            indexw.Close();
                        indexw = null;
                        if (filew != null)
                            filew.Close();
                        filew = null;
                    }
                    if (cancelSrc.Token.IsCancellationRequested)
                    {
                        Status(false, "", "Cancelled");
                        return;
                    }
                    if (sm.mailboxes != null)
                    {
                        Status(false, "", "Downloaded email to {0}", file);
                        foreach (Mailbox mailbox in sm.mailboxes)
                        {
                            if (mailbox.cnt != mailbox.start)
                            {
                                dldNotDone = true;
                                break;
                            }
                        }
                    }
                    else
                        dldNotDone = true;
                    if (dldNotDone)
                    {
                        Status(false, "", "Download is not complete, retrying...");
                        Thread.Sleep(5000);
                        await Start(attempts + 1);
                    }
                    else
                    {
                        TimeSpan span = DateTime.Now - start_time;
                        Status(false, "", "Download time {0} seconds", span.TotalSeconds);
                        File.Delete(indexFile);
                    }
                }
                else
                {
                    file = txtDownload.Text;
                    dir = Path.GetDirectoryName(txtDownload.Text);
                }
                if (dldNotDone == false && (bool)cbStatistics.IsChecked && !cancelSrc.Token.IsCancellationRequested)
                {
                    EmailStats stats = new EmailStats(cancelSrc.Token);
                    await stats.Start(dir, file, Status, delegate(double progress) { prBar.Value = progress; });
                    prBar.Value = 100;
                }
                if (cancelSrc.Token.IsCancellationRequested)
                    Status(false, "", "Cancelled");
            }
            catch (FailedLoginException)
            {
                Status(false, "", "Login failed, invalid user or password");
            }
            catch (FailedException)
            {
                Status(false, "", "Failed to download");
            }
            catch (SslFailedException)
            {
                Status(false, "", "Failed to establish secure connection");
            }
            catch (ValidationException ex)
            {
                Status(false,"",  ex.reason);
            }
            catch (CancelException)
            {
                Status(false, "", "Cancelled");
            }
            catch (Exception ex)
            {
                Status(false, "", ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                prBar.IsIndeterminate = false;
                prBar.Value = 0;
                btnBrowse.IsEnabled = true;
                btnStart.Content = "Start";
                if (filew != null)
                    filew.Close();
                if (indexw != null)
                    indexw.Close();
            }
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            if (cbDownload.IsChecked == true && cbResume.IsChecked == false)
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
                lblDownload.Content = "Choose directory where generated files are saved:";
            else
                lblDownload.Content = "Choose downloaded email archive:";
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
            cbResume.IsChecked = false;
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
                cbResume.Visibility = Visibility.Visible;
                txtDownload.Text = "";
            }
            lastDebug = e.Key;
            now = DateTime.Now;
        }

        private void cbResume_Checked(object sender, RoutedEventArgs e)
        {
            cbDownload.IsChecked = true;
            SetControls(true);
        }

        Dictionary<System.Windows.Controls.RadioButton,String> Providers()
        {
            return new Dictionary<System.Windows.Controls.RadioButton, string>()
                { { rb163, "163" }, {rbAol,"aol" }, {rbATT,"att" }, {rbComcast,"comcast" }, {rbCox, "cox" }, {rbGmail,"gmail"},
                  {rbGmx, "gmx" }, {rbHermes,"hermes" }, {rbiCloud, "iCloud" }, {rbInbox,"inbox" }, {rbMail,"mail" },
                  {rbOutlook, "outlook" }, {rbRambler,"rambler" }, {rbYahoo,"yahoo" }, {rbYandex,"yander" }, {rbYeah,"yeah" },
                  {rbZoho, "zoho" } };
        }

        private void SelectProvider()
        {
            if (rbOther.IsChecked == true)
            {
                cbAddr.Text = txtOtherImap.Text;
                txtPort.Text = txtOtherPort.Text;
                return;
            }
            Dictionary<System.Windows.Controls.RadioButton, String> providers = Providers();
            foreach (System.Windows.Controls.RadioButton rb in providers.Keys)
            {
                if (rb.IsChecked == true)
                {
                    cbAddr.SelectedValue = providers[rb];
                    txtPort.Text = "993";
                    break;
                }
            }
        }

        private void btnNextAcct_Click(object sender, RoutedEventArgs e)
        {
            tabControl.SelectedIndex = 1;
        }

        void SelectAccount()
        {
            txtUser.Text = txtUserAcct.Text;
            txtPassword.Password = txtPasswordAcct.Password;
        }

        private void btnNextAdv_Click(object sender, RoutedEventArgs e)
        {
            tabControl.SelectedIndex = 2;
        }

        private void rbOther_Checked(object sender, RoutedEventArgs e)
        {
            OtherControls(Visibility.Visible);
        }

        private void rbOther_Unchecked(object sender, RoutedEventArgs e)
        {
            OtherControls(Visibility.Hidden);
        }

        private void tabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (tabCtrlSelected == 2)
            {
                String addr = (string)cbAddr.Text;
                Dictionary<System.Windows.Controls.RadioButton, String> providers = Providers();
                bool found = false;
                if (addr != null && addr != "")
                {
                    foreach (System.Windows.Controls.RadioButton provider in providers.Keys)
                    {
                        if (providers[provider] == addr)
                        {
                            provider.IsChecked = true;
                            found = true;
                            OtherControls(Visibility.Hidden);
                            break;
                        }
                    }
                    if (found == false)
                    {
                        rbOther.IsChecked = true;
                        OtherControls(Visibility.Visible);
                        txtOtherImap.Text = addr;
                        txtOtherPort.Text = txtPort.Text;
                    }
                }

                txtUserAcct.Text = txtUser.Text;
                txtPasswordAcct.Password = txtPassword.Password;
            }
            else
            {
                SelectAccount();
                SelectProvider();
            }
            tabCtrlSelected = tabControl.SelectedIndex;
        }
    }
}
