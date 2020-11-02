using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Server.Application;

namespace Server.Middleware
{
    // You may need to install the Microsoft.AspNetCore.Http.Abstractions package into your project
    public class ExceptionMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly ILogger<ExceptionMiddleware> _logger;

        public ExceptionMiddleware(RequestDelegate next, ILogger<ExceptionMiddleware> logger)
        {
            _next = next;
            _logger = logger;
        }

        public async Task Invoke(HttpContext httpContext)
        {

            try
            {
                await _next(httpContext);
            }
            catch (Exception e)
            {
                await HandleExceptionAsync(httpContext, e, _logger);
            }
        }

        private async Task HandleExceptionAsync(HttpContext httpContext, Exception e, ILogger<ExceptionMiddleware> logger)
        {
            string response = string.Empty;

            switch (e)
            {
                case RESTfulException restfulException:
                    logger.LogError(restfulException, "RESET ful Error");
                    httpContext.Response.StatusCode = (int)restfulException.StatusCode;
                    response =  JsonConvert.SerializeObject(new
                    {
                        Code = (int)restfulException.StatusCode,
                        restfulException.Errors
                    });
                    break;

                case Exception ex:
                    logger.LogError(ex, "SERVER Error");
                    httpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                    response = ex.Message;
                    break;

            }

            httpContext.Response.ContentType = "application/json";
            if (string.IsNullOrEmpty(response)) return;

            await httpContext.Response.WriteAsync(response);
        }
    }

    // Extension method used to add the middleware to the HTTP request pipeline.
    public static class ExceptionMiddlewareExtensions
    {
        public static IApplicationBuilder UseExceptionMiddleware(this IApplicationBuilder builder)
        {
            return builder.UseMiddleware<ExceptionMiddleware>();
        }
    }
}
