using Microsoft.AspNetCore.Identity;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace API.Models
{
    [Table("Tb_M_UserRole")]
    public class UserRole : IdentityUserRole<string>
    {
        //[ForeignKey("User")]
        //public override string UserId { get; set; }

        //[ForeignKey("Role")]
        //public override string RoleId { get; set; }
        
        public User User { get; set; }

        public Role Role { get; set; }
    }
}
