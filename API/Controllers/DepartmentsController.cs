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
    public class DepartmentsController : BaseController<Department, DepartmentRepo>
    {
        readonly DepartmentRepo _departRepo;
        public DepartmentsController(DepartmentRepo departmentRepo) : base(departmentRepo)
        {
            _departRepo = departmentRepo;
        }

        [HttpPut("{id}")]
        public async Task<ActionResult<int>> Update(int id, Department entity)
        {
            var getId = await _departRepo.GetID(id);
            getId.Name = entity.Name;
            var data = await _departRepo.Update(getId);
            if (data.Equals(null))
            {
                return BadRequest("Something Wrong! Please check again");
            }
            return data;
        }
    }
}
