using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Net.Security;
using System.IO;

namespace email_proc
{
    public enum ConnectStatus { Ok, Failed }
    class ImapConnect
    {
        String host {get; set; }
        int port { get; set; }
        TcpClient tcpc { get; set; }
        public Stream stream { get; private set; }

        public ImapConnect(String host, int port)
        {
            this.host = host;
            this.port = port;
            tcpc = null;
            stream = null;
        }

        public async Task<ConnectStatus> Connect()
        {
            try
            {
                tcpc = new TcpClient();
                tcpc.NoDelay = true;
                await tcpc.ConnectAsync(host, port);
                stream = tcpc.GetStream();
                return ConnectStatus.Ok;
            }
            catch (Exception ex)
            {
                return ConnectStatus.Failed;
            }
        }

        public async Task<ConnectStatus> ConnectTls ()
        {
            try {
                SslStream ssl = new SslStream(tcpc.GetStream(), false, CertificateValidationCallBack);
                await ssl.AuthenticateAsClientAsync(host);
                stream = ssl;
                return ConnectStatus.Ok;
            }
            catch (Exception ex)
            {
                return ConnectStatus.Failed;
            }
        }

        public void Close()
        {
            if (tcpc != null)
                tcpc.Close();
        }

        private static bool CertificateValidationCallBack(
         object sender,
         System.Security.Cryptography.X509Certificates.X509Certificate certificate,
         System.Security.Cryptography.X509Certificates.X509Chain chain,
         System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }
    }
}
