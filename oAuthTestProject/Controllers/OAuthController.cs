using Microsoft.Graph;
using oAuthTestProject.Providers;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace oAuthTestProject.Controllers
{
    public class OAuthController : ApiController
    {
        [Route("PruebaOAuth")]
        [HttpPost]
        public async Task<HttpResponseMessage> GraphTestAsync()
        {
            MicrosoftGraphProvider _microsoftGraphProvider = new MicrosoftGraphProvider();
            OAuthViewModel objresult = new OAuthViewModel();
            string emailOrObjectId = ConfigurationManager.AppSettings["EmailOrObjectId"];
            string fileNameOrObjectId = ConfigurationManager.AppSettings["FileNameOrObjectId"];

            //conseguir usuario
            var ResponseData = await _microsoftGraphProvider.GetIdByEmail(emailOrObjectId);

            if (ResponseData.Status.Code != 200)
            {
                objresult.OnboardingOAuthData = new { ServiceResponse = false };
                objresult.HttpResponse = new { Code = ResponseData.Status.Code, Message = "Error usuario: " + ResponseData.Status.Message };
                return Request.CreateResponse(HttpStatusCode.OK, objresult);
            }

            //descargar archivo .docx
            var ResponseDataFile = await _microsoftGraphProvider.GetFileStreamByUserId(ResponseData, emailOrObjectId, fileNameOrObjectId);

            if (ResponseDataFile.Status.Code != 200)
            {
                objresult.OnboardingOAuthData = new { ServiceResponse = false };
                objresult.HttpResponse = new { Code = ResponseDataFile.Status.Code, Message = "Error descarga: " + ResponseDataFile.Status.Message };
                return Request.CreateResponse(HttpStatusCode.OK, objresult);
            }

            //Subir archivo .docx
            var ResponseDataUpload = await _microsoftGraphProvider.UploadFileOneDrive(ResponseDataFile, emailOrObjectId, fileNameOrObjectId);

            if (ResponseDataUpload.Status.Code != 200)
            {
                objresult.OnboardingOAuthData = new { ServiceResponse = false };
                objresult.HttpResponse = new { Code = ResponseDataFile.Status.Code, Message = "Error subida: " + ResponseDataFile.Status.Message };
                return Request.CreateResponse(HttpStatusCode.OK, objresult);
            }

            //Descargar como .pdf
            var PdfResponseFile = await _microsoftGraphProvider.ConvertFileToPdf(ResponseDataUpload, emailOrObjectId, fileNameOrObjectId);

            if (PdfResponseFile.Status.Code != 200)
            {
                objresult.OnboardingOAuthData = new { ServiceResponse = false };
                objresult.HttpResponse = new { Code = ResponseDataUpload.Status.Code, Message = "Error descarga pdf: " + ResponseDataUpload.Status.Message };
                return Request.CreateResponse(HttpStatusCode.OK, objresult);
            }

            objresult.OnboardingOAuthData = new { ServiceResponse = true };
            objresult.HttpResponse = new { Code = ResponseData.Status.Code, Message = "Ok" };
            return Request.CreateResponse(HttpStatusCode.OK, objresult);
        }

        public class OAuthViewModel
        {
            public Object HttpResponse { get; set; }
            public object OnboardingOAuthData { get; set; }
        }
    }
}

