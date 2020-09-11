using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using API.Context;
using API.ViewModels;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ChartsController : ControllerBase
    {
        private readonly MyContext _context;
        public ChartsController(MyContext myContext)
        {
            _context = myContext;
        }
        // GET api/values
        [HttpGet]
        [Route("pie")]
        public async Task<List<PieChartVM>> GetPie()
        {
            //var user = new UserVM();
            //var getData = await _context.Divisions
            //                    .Join(
            //                        _context.Departments,
            //                        di => di.DepartmentId,
            //                        de => de.Id,
            //                        (di,de) => new { Divisions = di, Departments = de })
            //                    .Where(x => x.Divisions.isDelete == false)
            //                    .ToListAsync();

            //var getData = await _context.Divisions.Include("Department").Where(x => x.isDelete == false).ToListAsync();
            //var data = await _context.Divisions
            //                .Join(_context.Departments, 
            //                        di => di.DepartmentId, 
            //                        de => de.Id, 
            //                        (di, de) => new { 
            //                            Divisions = di, Departments = de 
            //                        }).GroupBy(q => q.Departments.Name).Select(q => new
            //                        {
            //                            GroupId = q.Key,
            //                            Count = q.Count()
            //                        }).ToListAsync();

            var data1 = await _context.Divisions.Include("Department")
                            .Where(x => x.isDelete == false)
                            .GroupBy(q => q.Department.Name)
                            .Select(q => new PieChartVM
                             {
                                    DepartmentName = q.Key,
                                    total = q.Count()
                             }).ToListAsync();
            return data1;
        }
    }
}
