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
    public class RowsController : ControllerBase
    {
        // GET: api/Rows
        [HttpGet]
        public IEnumerable<Row> Get()
        {
            Xl xl = new Xl();
            return xl.GetRows();
        }

        // GET: api/Rows/5
        [HttpGet("{id}")]
        public Row Get(int id)
        {
            Xl xl = new Xl();
            return xl.GetRow(id);
        }
    }
}
