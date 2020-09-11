using API.Context;
using API.Models;
using API.ViewModels;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace API.Repository.Data
{
    public class DivisionRepo : GeneralRepo<Division, MyContext>
    {
        readonly MyContext _context;
        IConfiguration _configuration;
        public DivisionRepo(MyContext context, IConfiguration config) : base(context)
        {
            _context = context;
            _configuration = config;
        }

        public override async Task<List<Division>> GetAll()
        {
            List<DivisionVM> list = new List<DivisionVM>();
            var data = await _context.Divisions.Include("Department").Where(x => x.isDelete == false).ToListAsync();
            if (data.Count == 0)
            {
                return null;
            }
            return data;
        }
        public override async Task<Division> GetID(int Id)
        {
            var data = await _context.Divisions.Include("Department").SingleOrDefaultAsync(x => x.Id == Id && x.isDelete == false);
            if (!data.Equals(0))
            {
                return data;
            }
            return null;
        }
    }
}
