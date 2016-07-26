using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;

namespace ExcelToJson.App_Start
{
    public class CORSPreflightModule : IHttpModule
    {
        private const string OPTIONSMETHOD = "OPTIONS";
        private const string ORIGINHEADER = "ORIGIN";
        private const string ALLOWEDORIGIN = "https://wipronvs1.sharepoint.com";
        void IHttpModule.Dispose()
        {

        }
        void IHttpModule.Init(HttpApplication context)
        {
            context.PreSendRequestHeaders += (sender, e) =>
            {
                var response = context.Response;

                if (context.Request.Headers[ORIGINHEADER] == ALLOWEDORIGIN)
                {
                    response.Headers.Add("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
                    response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
                }
                if (context.Request.HttpMethod.ToUpperInvariant() == OPTIONSMETHOD && context.Request.Headers[ORIGINHEADER] == ALLOWEDORIGIN)
                {
                    response.Headers.Clear();
                    response.Headers.Add("Access-Control-Allow-Methods", "GET,POST,PUT,DELETE,OPTIONS");
                    response.Headers.Add("Access-Control-Allow-Origin", "https://yourspodomain.sharepoint.com");
                    response.Headers.Add("Access-Control-Allow-Credentials", "true");
                    response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
                    response.Clear();
                    response.StatusCode = (int)HttpStatusCode.OK;
                }
            };

        }

    }

}