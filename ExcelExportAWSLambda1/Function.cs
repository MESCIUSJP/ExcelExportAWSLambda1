using Amazon.Lambda.APIGatewayEvents;
using Amazon.Lambda.Core;
using GrapeCity.Documents.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace ExcelExportAWSLambda1
{
    public class Function
    {

        public APIGatewayProxyResponse FunctionHandler(APIGatewayProxyRequest input, ILambdaContext context)
        {
            APIGatewayProxyResponse response;

            string queryString;
            input.QueryStringParameters.TryGetValue("name", out queryString);

            string Message = string.IsNullOrEmpty(queryString)
                ? "Hello, World!!"
                : $"Hello, {queryString}!!";

            //Workbook.SetLicenseKey("");

            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Range["A1"].Value = Message;

            var base64String = "";

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFileFormat.Xlsx);
                base64String = Convert.ToBase64String(ms.ToArray());
            }

            response = new APIGatewayProxyResponse
            {
                StatusCode = (int)HttpStatusCode.OK,
                Body = base64String,
                IsBase64Encoded = true,
                Headers = new Dictionary<string, string> {
                    { "Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                    { "Content-Disposition", "attachment; filename=Result.xlsx"},
                }
            };

            return response;
        }
    }
}
