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
    public class RowsController : ControllerBase
    {
        private Xl _xl;
        public RowsController(Xl xl)
        {
            _xl = xl;
        }
        // GET: api/Rows
        [HttpGet]
        [Route("api/Rows")]
        public IEnumerable<Row> Get()
        {
            return _xl.GetRows();
        }

        // GET: api/Rows/5
        [HttpGet("{id}")]
        [Route("api/Rows/{id}")]
        public Row Get(int id)
        {
            return _xl.GetRow(id);
        }
    }
}
