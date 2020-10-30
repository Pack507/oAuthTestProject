using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace oAuthTestProject.Providers
{
    public interface IGraphProvider
    {
        Task<OAuthUserModel> GetIdByEmail(string email);
        Task<OAuthUserModel> GetFileStreamByUserId(OAuthUserModel oAuth, string EmailOrObjectId, string FileNameOrObjectId);
        Task<OAuthUserModel> UploadFileOneDrive(OAuthUserModel oAuth, string emailOrObjectId, string fileNameOrObjectId);
        Task<OAuthUserModel> ConvertFileToPdf(OAuthUserModel oAuth, string emailOrObjectId, string fileNameOrObjectId);
    }
    public class MicrosoftGraphProvider : IGraphProvider
    {
        private string _tenantId { get; set; }
        private string _clientId { get; set; }
        private string _clientSecret { get; set; }

        public MicrosoftGraphProvider()
        {
            _tenantId = ConfigurationManager.AppSettings["TenantId"];
            _clientId = ConfigurationManager.AppSettings["ClientId"];
            _clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
        }

        public async Task<OAuthUserModel> GetIdByEmail(string EmailOrObjectId)
        {
            OAuthUserModel model = new OAuthUserModel();
            try
            {
                var confidentialClient = ConfidentialClientApplicationBuilder
                .Create(_clientId)
                .WithTenantId(_tenantId)
                .WithClientSecret(_clientSecret)
                .Build();
                ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClient);
                GraphServiceClient graphClient = new GraphServiceClient(authProvider);
                var user = await graphClient.Users[EmailOrObjectId].Request().GetAsync();

                model.User = user;
                model.AuthProvider = authProvider;
                model.GraphService = graphClient;
                model.Status.Code = 200;
                model.Status.Message = "OK";

                return model;
            }
            catch (Exception e)
            {
                model.Status.Code = 406;
                model.Status.Message = e.Message;
                return model;
            }
        }

        public async Task<OAuthUserModel> GetFileStreamByUserId(OAuthUserModel oAuth, string emailOrObjectId, string fileNameOrObjectId)
        {
            try
            {
                Stream templateFileStream = await oAuth.GraphService.Users[emailOrObjectId].Drive.Items[fileNameOrObjectId].Content.Request().GetAsync();

                oAuth.TemplateStream = templateFileStream;
                oAuth.Status.Code = 200;
                oAuth.Status.Message = "OK";

                return oAuth;
            }
            catch (ServiceException e)
            {
                oAuth.Status.Code = 406;
                oAuth.Status.Message = e.Message;
                return oAuth;
            }
        }

        public async Task<OAuthUserModel> UploadFileOneDrive(OAuthUserModel oAuth, string emailOrObjectId, string fileNameOrObjectId)
        {
            try
            {
                string _newFileName = $"Prueba Onboarding-{Guid.NewGuid()}.docx";

                // where you want to save the file, with name
                var item = $"/Contratos/" + _newFileName;

                var uploadSession = await GetUploadSession(oAuth, item, emailOrObjectId);

                var maxChunkSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
                var provider = new ChunkedUploadProvider(uploadSession, oAuth.GraphService, oAuth.TemplateStream, maxChunkSize);

                // Setup the chunk request necessities
                var chunkRequests = provider.GetUploadChunkRequests();
                var readBuffer = new byte[maxChunkSize];
                var trackedExceptions = new List<Exception>();
                DriveItem itemResult = null;

                //upload the chunks
                foreach (var request in chunkRequests)
                {
                    // Do your updates here: update progress bar, etc.
                    // ...
                    // Send chunk request
                    var result = await provider.GetChunkRequestResponseAsync(request, readBuffer, trackedExceptions);

                    if (result.UploadSucceeded)
                    {
                        itemResult = result.ItemResponse;
                    }
                }

                oAuth.Status.Code = 200;
                oAuth.Status.Message = "OK";

                return oAuth;
            }
            catch (Exception e)
            {
                oAuth.Status.Code = 406;
                oAuth.Status.Message = e.Message;
                return oAuth;
            }
        }
        public async Task<OAuthUserModel> ConvertFileToPdf(OAuthUserModel oAuth, string emailOrObjectId, string fileNameOrObjectId)
        {
            try
            {
                var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("format", "pdf")
                };

                var stream = await oAuth.GraphService.Users[emailOrObjectId].Drive.Items[fileNameOrObjectId].Content.Request(queryOptions).GetAsync();

                oAuth.PdfStream = stream;
                oAuth.Status.Code = 200;
                oAuth.Status.Message = "Ok";

                return oAuth;
            }
            catch (Exception e)
            {
                oAuth.Status.Code = 406;
                oAuth.Status.Message = e.Message;
                return oAuth;
            }
        }
        public async Task<UploadSession> GetUploadSession(OAuthUserModel oAuth, string item, string user)
        {
            return await oAuth.GraphService.Users[user].Drive.Root.ItemWithPath(item).CreateUploadSession().Request().PostAsync();
        }
    }

    public class OAuthUserModel
    {
        public User User { get; set; }
        public ClientCredentialProvider AuthProvider { get; set; }
        public GraphServiceClient GraphService { get; set; }
        public Stream TemplateStream { get; set; }
        public Stream PdfStream { get; set; }
        public OAuthStatusModel Status { get; set; }
        public OAuthUserModel()
        {
            User = new User();
            GraphService = new GraphServiceClient(AuthProvider);
            Status = new OAuthStatusModel();
        }
    }
    public class OAuthStatusModel
    {
        public int Code { get; set; }
        public string Message { get; set; }
    }
}