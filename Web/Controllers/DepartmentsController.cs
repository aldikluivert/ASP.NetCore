using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using API.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace Web.Controllers
{
    public class DepartmentsController : Controller
    {
        readonly HttpClient client = new HttpClient
        {
            BaseAddress = new Uri("https://localhost:44374/api/")
        };

        public IActionResult Index()
        {
            if (HttpContext.Session.GetString("lvl") == "Admin")
            {
                return View("~/Views/Dashboard/Department.cshtml");
            }
            return Redirect("/notfound");
        }

        public IActionResult LoadDepart()
        {
            IEnumerable<Department> departments = null;
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);
            var resTask = client.GetAsync("departments");
            resTask.Wait();

            var result = resTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<List<Department>>();
                readTask.Wait();
                departments = readTask.Result;
            }
            else
            {
                departments = Enumerable.Empty<Department>();
                ModelState.AddModelError(string.Empty, "Server Error try after sometimes.");
            }
            return Json(departments);

        }

        public IActionResult GetById(int Id)
        {
            Department departments = null;
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);
            var resTask = client.GetAsync("departments/" + Id);
            resTask.Wait();

            var result = resTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var json = JsonConvert.DeserializeObject(result.Content.ReadAsStringAsync().Result).ToString();
                departments = JsonConvert.DeserializeObject<Department>(json);
            }
            else
            {
                ModelState.AddModelError(string.Empty, "Server Error.");
            }
            return Json(departments);
        }

        public IActionResult InsertOrUpdate(Department departments, int id)
        {
            try
            {
                var json = JsonConvert.SerializeObject(departments);
                var buffer = System.Text.Encoding.UTF8.GetBytes(json);
                var byteContent = new ByteArrayContent(buffer);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                var token = HttpContext.Session.GetString("token");
                client.DefaultRequestHeaders.Add("Authorization", token);
                if (departments.Id == 0)
                {
                    var result = client.PostAsync("departments", byteContent).Result;
                    return Json(result);
                }
                else if (departments.Id == id)
                {
                    var result = client.PutAsync("departments/" + id, byteContent).Result;
                    return Json(result);
                }

                return Json(404);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IActionResult Delete(int id)
        {
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);
            var result = client.DeleteAsync("departments/" + id).Result;
            return Json(result);
        }
    }
}
