using Email.Office365;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace Email
{
    public class EmailHelper
    {
        private static readonly Regex[] _htmlReplaces = new[] {
            new Regex(@"<script\b[^<]*(?:(?!</script>)<[^<]*)*</script>", RegexOptions.Compiled | RegexOptions.Singleline),
            new Regex(@"<style\b[^<]*(?:(?!</style>)<[^<]*)*</style>", RegexOptions.Compiled | RegexOptions.Singleline),
            new Regex(@"<[^>]*>", RegexOptions.Compiled),
            new Regex(@" +", RegexOptions.Compiled)
        };

        private readonly string _user;
        private readonly string _password;
        private readonly string _emailUser;

        private readonly int _port;
        private readonly string _server;

        private readonly ExchangeService _exchangeService;

        public List<EmailModel> EmailsInBox = new List<EmailModel>();

        public EmailModel PrimeiroEmail { get; private set; }

        public bool GetErro { get; private set; }

        public EmailHelper()
        {
            GetErro = true;
        }

        public EmailHelper( string user, string password, 
                            string emailUser, string url, 
                            string domain, int port, 
                            string clientSecret = null, string clientId = null, string tenantId = null)
        {
            var uri = new Uri(url);

            _user = user;
            _password = password;
            _emailUser = emailUser;
            _port = port;
            _server = domain;

            GetErro = false;

            try
            {
                if (clientSecret != null)
                {
                    var accessToken = ApiHandler.GetAccessTokenAsync(clientSecret, clientId, tenantId).Result;
                    if (accessToken.access_token == null)
                        throw new Exception("Não foi possível obter o access token, não existe ou é nulo.");

                    var exchange = new ExchangeService()
                    {
                        Url = uri,
                        KeepAlive = true,
                        UseDefaultCredentials = false,
                        Credentials = new OAuthCredentials(accessToken.access_token),
                        ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, _emailUser),
                        TraceEnablePrettyPrinting = true,
                        Timeout = 120000
                    };

                    exchange.HttpHeaders.Add("X-AnchorMailbox", _emailUser);

                    _exchangeService = exchange ?? throw new Exception("Não foi possivel conectar com o Office 365. ");
                }
                else
                {
                    var exchange = new ExchangeService(ExchangeVersion.Exchange2010)
                    {
                        Url = uri,
                        KeepAlive = true,
                        UseDefaultCredentials = false,
                        Credentials = new WebCredentials(user, password),
                        TraceEnablePrettyPrinting = true,
                        Timeout = 120000
                    };

                    _exchangeService = exchange ?? throw new Exception("Não foi possivel conectar com o Exchange. ");

                    _exchangeService.AutodiscoverUrl(_emailUser);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public bool CompararMesEAno(DateTime data1, DateTime data2)
        {
            return data1.Month == data2.Month && data1.Year == data2.Year;
        }

        public void EnviarExpiracaoClientSecret(DateTime dataExpiracao, string emailsValidos)
        {
            var emails = emailsValidos.Split(',');

            string mensagem = $" Envio automático, Por Favor não responder. \n" +
            $"A data de expiração da chave secreta é dia {dataExpiracao:dd/MM/yyyy}. " +
            $"Considere atualizar a chave secreta desse Email de serviço no portal.azure.com " +
            $" antes que o serviço fique indiponível no aeroporto.";

            foreach (var item in emails)
            {
                if (!string.IsNullOrEmpty(item))
                    EnviarEmail(mensagem, "Expiração de Chave Secreta do Email de Serviço", item);
            }
        }

        public void MarcarEmailComoLido(string id)
        {
            try
            {
                var email = EmailMessage.Bind(_exchangeService, new ItemId(id));

                if (email == null)
                    throw new Exception("Email não encontrado");

                email.IsRead = true;
                email.IsReadReceiptRequested = true;
            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possível marcar o email como lido. " + ex.Message);
            }
        }

        public void ObterTodosEmailsInbox()
        {
            try
            {
                if (_exchangeService != null)
                {
                    //ordenação by ReceivedTime
                    var view = new ItemView(100);
                    view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

                    var emails = _exchangeService.FindItems(WellKnownFolderName.Inbox, view);
                    
                    //'Enquanto existir mensagens'
                    if (emails.Count() > 0)
                    {
                        foreach (Item item in emails)
                        {
                            EmailMessage email = EmailMessage.Bind(_exchangeService, item.Id);

                            var emailId = email.Id.UniqueId;

                            if (email.ItemClass.ToUpper() == "IPM.NOTE")
                            {
                                var subject = email.Subject;
                                if (subject.Contains("SOLICITACAO SCORE"))
                                {
                                    email.Load(PropertySet.FirstClassProperties);

                                    var emailOrigem = email.From.Address;

                                    if (email.Body.BodyType == BodyType.HTML)
                                    {
                                        var html = HtmlToPlainText(email.Body.Text);
                                        AdicionarEmail(emailId, subject, emailOrigem, html);
                                    }
                                    else
                                    {
                                        var body = email.Body.Text;
                                        AdicionarEmail(emailId, subject, emailOrigem, body);
                                    }
                                }
                                else
                                {
                                    var itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties)
                                    {
                                        RequestedBodyType = BodyType.Text
                                    };
                                    email.Load(itempropertyset);

                                    var emailOrigem = email.From.Address;
                                    var body = email.Body.Text;

                                    AdicionarEmail(emailId, subject, emailOrigem, body);
                                }
                            }
                            else
                            {
                                RemoverEmailInbox(emailId);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GetErro = true;
                throw new Exception("Não foi possível conectar para obter os emails da caixa de entrada." + ex.Message);
            }
        }

        private void AdicionarEmail(string emailId, string subject, string emailOrigem, string body)
        {
            EmailsInBox.Add(new EmailModel
            {
                Id = emailId,
                From = emailOrigem,
                Subject = subject,
                BodyText = body,
            });
        }

        public void RemoverTodosEmailsDeletedItems()
        {
            try
            {
                var emails = _exchangeService.FindItems(WellKnownFolderName.DeletedItems, new ItemView(100));

                if (emails.Count() > 0)
                    _exchangeService.DeleteItems(emails.Select(email => email.Id), DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possível remover os emails da lixeira. " + ex.Message);
            }
        }

        public void RemoverEmailInbox(string id)
        {
            try
            {
                var email = EmailMessage.Bind(_exchangeService, new ItemId(id));

                if (email == null)
                    throw new Exception("Email não encontrado");

                email.Delete(DeleteMode.MoveToDeletedItems);
            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possível remover o email para lixeira. " + ex.Message);
            }
        }

        public void EnviarEmail(string mensagem, string assunto, string emailTo)
        {
            try
            {
                var message = new EmailMessage(_exchangeService);

                if (string.IsNullOrEmpty(emailTo))
                    throw new Exception("É necessário um email para o envio da mensagem.");

                message.ToRecipients.Add(emailTo);

                message.Subject = assunto;

                message.Body = new MessageBody(BodyType.Text, mensagem);

                message.Send();
            }
            catch (Exception ex)
            {
                GetErro = true;
                throw new Exception("O Exchange Service não conseguiu enviar o email. " + ex.Message);
            }
        }

        private void Enviar(MailMessage message)
        {
            try
            {
                using (var client = new SmtpClient())
                {
                    client.Host = _server;

                    var basicCredential = new NetworkCredential(_user, _password);
                    client.UseDefaultCredentials = false;
                    client.Credentials = basicCredential;

                    client.Send(message);
                }
            }
            catch (Exception ex)
            {
                GetErro = true;
                throw new Exception(ex.Message);
            }
        }

        private string HtmlToPlainText(string html)
        {
            foreach (var r in _htmlReplaces)
            {
                html = r.Replace(html, "");
            }
            var lines = html
                .Split(new[] { '\r', '\n' })
                .Select(_ => WebUtility.HtmlDecode(_.Trim()))
                .Where(_ => _.Length > 0)
                .ToArray();

            /* Replace do Caracter "\u00A0" para espaço em branco
             * para que o Aixs consiga ler o E-mail */
            return Regex.Replace(string.Join("\n", lines), @"\u00A0", " ");
        }
    }
}