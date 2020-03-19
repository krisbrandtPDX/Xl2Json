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
    public class BooksController : ControllerBase
    {
        private readonly Xl _xl;
        public BooksController(Xl xl)
        {
            _xl = xl;
        }

        // GET: api/Books
        [HttpGet]
        [Route("api/Books")]
        public IEnumerable<Book> Get()
        {
            return _xl.GetBooks();
        }

        // GET: api/Books/5
        [HttpGet]
        [Route("api/Books/{id}")]
        public Book Get(int id)
        {
            return _xl.GetBook(id);
        }

    }
}
