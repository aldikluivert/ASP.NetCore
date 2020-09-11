using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using API.Base;
using API.Models;
using API.Repository.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [Authorize(AuthenticationSchemes = "Bearer")]
    [Route("api/[controller]")]
    [ApiController]
    public class DivisionsController : BaseController<Division, DivisionRepo> 
    {
        readonly DivisionRepo _divisionRepo;
        public DivisionsController(DivisionRepo divisionRepo) : base(divisionRepo)
        {
            _divisionRepo = divisionRepo;
        }

        [HttpPut("{id}")]
        public async Task<ActionResult<int>> Update(int id, Division entity)
        {
            var getId = await _divisionRepo.GetID(id);
            getId.Name = entity.Name;
            getId.DepartmentId = entity.DepartmentId;
            var data = await _divisionRepo.Update(getId);
            if (data.Equals(null))
            {
                return BadRequest("Something Wrong! Please check again");
            }
            return data;
        }
    }
}
