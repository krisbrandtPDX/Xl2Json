using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Xl2Json.Models;

namespace Xl2Json.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SheetsController : ControllerBase
    {
        // GET: api/Sheets
        [HttpGet]
        public IEnumerable<Sheet> Get()
        {
            Xl xl = new Xl();
            return xl.GetSheets();
        }

        // GET: api/Sheets/5
        [HttpGet("{id}")]
        public Sheet Get(int id)
        {
            Xl xl = new Xl();
            return xl.GetSheet(id);
        }
    }
}
