using Microsoft.AspNetCore.Identity;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

namespace API.Models
{
    [Table("TB_M_User")]
    public class User : IdentityUser
    {
        public ICollection<UserRole> userRoles { get; set; }
        public Employee Employee { get; set; }
    }
}
