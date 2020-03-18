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
    public class BooksController : ControllerBase
    {
        // GET: api/Books
        [HttpGet]
        public IEnumerable<Book> Get()
        {
            Xl xl = new Xl();
            return xl.GetBooks();
        }

        // GET: api/Books/5
        [HttpGet("{id}")]
        public Book Get(int id)
        {
            Xl xl = new Xl();
            return xl.GetBook(id);
        }
    }
}
