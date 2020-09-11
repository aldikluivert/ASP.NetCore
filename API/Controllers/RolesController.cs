using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using API.Context;
using API.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RolesController : ControllerBase
    {
        private readonly MyContext _context;
        private readonly UserManager<Role> _userManager;
        public RolesController(MyContext myContext, UserManager<Role> userManager)
        {
            _context = myContext;
            _userManager = userManager;
        }

    }
}
