using AshfordSync.Entities;
using AshfordSync.Interfaces;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AshfordSync.Service
{
    class ExcelService : IExcelService
    {
        static string[] Scopes = { GmailService.Scope.GmailReadonly, GmailService.Scope.GmailModify };
        static string ApplicationName = "AshfordSync";
        private readonly ILogger<ExcelService> _logger;
        private readonly IReadInventoryService _service;

        public ExcelService(ILogger<ExcelService> logger, IReadInventoryService service)
        {
            _logger = logger;
            _service = service;
        }   

        public async Task ProcessExcel()
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            var inboxlistRequest = service.Users.Messages.List("me");
            inboxlistRequest.LabelIds = "INBOX";
            inboxlistRequest.IncludeSpamTrash = false;

            var emailListResponse = inboxlistRequest.Execute();
            if (emailListResponse.Messages != null)
            {
                foreach (var mail in emailListResponse.Messages)
                {
                    var mailId = mail.Id;
                    var threadId = mail.ThreadId;
                    string subject = "";
                    string from = "";
                    try
                    {
                        Message message = service.Users.Messages.Get("me", mailId).Execute();

                        foreach (MessagePartHeader header in message.Payload.Headers)
                        {
                            Console.WriteLine(header.Name);
                            _logger.LogInformation(header.Name);
                            Console.WriteLine(header.Value);
                            _logger.LogInformation(header.Value);
                            if (header.Name.Equals("Subject"))
                            {
                                subject = header.Value;
                            }
                            if (header.Name.Equals("From"))
                            {
                                from = header.Value;
                            }
                        }
                        Console.WriteLine("Valid Domain " + jsonParamModel.validDomain);
                        _logger.LogInformation("Valid Domain " + jsonParamModel.validDomain);
                        Console.WriteLine("From" + from);
                        _logger.LogInformation("From" + from);
                        if (from.Contains(jsonParamModel.validDomain))
                        {
                                                        
                            IList<MessagePart> parts = message.Payload.Parts;
                            foreach (MessagePart part in parts)
                            {
                                if (!String.IsNullOrEmpty(part.Filename))
                                {
                                    String attId = part.Body.AttachmentId;
                                    MessagePartBody attachPart = service.Users.Messages.Attachments.Get("me", mailId, attId).Execute();

                                    
                                    String attachData = attachPart.Data.Replace('-', '+');
                                    attachData = attachData.Replace('_', '/');

                                    byte[] data = Convert.FromBase64String(attachData);
                                    File.WriteAllBytes(Path.Combine(".\\Inbox\\", part.Filename), data);
                                    service.Users.Messages.Trash("me", mailId).Execute();



                                    _logger.LogInformation("Start Processing File:" + part.Filename);

                                    int supplierId = jsonParamModel.supplierId;
                                    
                                    if (part.Filename.Contains("inventory", System.StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        await _service.ReadInventoryAsync(supplierId, part.Filename);
                                    }
                                    else if (part.Filename.Contains("ship", System.StringComparison.CurrentCultureIgnoreCase) || part.Filename.Contains("fulfill", System.StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        await _service.ReadShipConfirmAsync(supplierId, part.Filename);
                                    }
                                    else if (part.Filename.Contains("rma", System.StringComparison.CurrentCultureIgnoreCase) || part.Filename.Contains("return", System.StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        await _service.ReadRMAAsync(supplierId, part.Filename);
                                    }

                                    _logger.LogInformation("End Processing File:" + part.Filename);

                                }
                            }
                        }
                        service.Users.Messages.Trash("me", mailId).Execute();

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("An error occurred: " + e.Message);
                        _logger.LogError("An error occurred: " + e.Message);
                    }

                }
            }
        }
    }
}
