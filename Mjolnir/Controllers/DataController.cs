using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace Mjolnir.Controllers
{
    [Produces("application/json")]
    [Route("api/data")]
    public class DataController : Controller
    {
       [HttpGet]
       public string Get()
       {
            return "Data info";
       }
    }
}