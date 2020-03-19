using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Xl2Json.Models;

namespace Xl2Json.Controllers
{
    [ApiController]
    public class SheetsController : ControllerBase
    {
        private Xl _xl;
        public SheetsController(Xl xl)
        {
            _xl = xl;
        }
        // GET: api/Sheets
        [HttpGet]
        [Route("api/Sheets")]
        public IEnumerable<Sheet> Get()
        {
            return _xl.GetSheets();
        }

        // GET: api/Sheets/5
        [HttpGet("{id}")]
        [Route("api/Sheets/{id}")]
        public Sheet Get(int id)
        {
            return _xl.GetSheet(id);
        }
    }
}
