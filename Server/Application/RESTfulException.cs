using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace Server.Application
{
    public class RESTfulException : Exception
    { 
        public HttpStatusCode StatusCode { get; }

        public object Errors { get; }

        public RESTfulException(HttpStatusCode statusCode, object errors)
        {
            StatusCode = statusCode;
            Errors = errors;
        }
    }
}
