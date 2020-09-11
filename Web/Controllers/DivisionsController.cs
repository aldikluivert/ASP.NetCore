using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using API.Models;
using API.ViewModels;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Web.Report;

namespace Web.Controllers
{
    public class DivisionsController : Controller
    {
        readonly HttpClient client = new HttpClient
        {
            BaseAddress = new Uri("https://localhost:44374/api/")
        };

        public IActionResult Index()
        {
            if (HttpContext.Session.GetString("lvl") == "Admin")
            {
                return View("~/Views/Dashboard/Division.cshtml");
            }
            return Redirect("/notfound");
        }

        public IActionResult LoadDiv()
        {
            IEnumerable<Division> division = null;
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);
            var resTask = client.GetAsync("divisions");
            resTask.Wait();

            var result = resTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<List<Division>>();
                readTask.Wait();
                division = readTask.Result;
            }
            else
            {
                division = Enumerable.Empty<Division>();
                ModelState.AddModelError(string.Empty, "Server Error try after sometimes.");
            }
            return Json(division);

        }

        public IActionResult GetById(int Id)
        {
            Division division = null;
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);
            var resTask = client.GetAsync("divisions/" + Id);
            resTask.Wait();

            var result = resTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var json = JsonConvert.DeserializeObject(result.Content.ReadAsStringAsync().Result).ToString();
                division = JsonConvert.DeserializeObject<Division>(json);
            }
            else
            {
                ModelState.AddModelError(string.Empty, "Server Error.");
            }
            return Json(division);
        }

        public IActionResult InsertOrUpdate(Division data, int id)
        {
            try
            {
                var json = JsonConvert.SerializeObject(data);
                var buffer = System.Text.Encoding.UTF8.GetBytes(json);
                var byteContent = new ByteArrayContent(buffer);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                var token = HttpContext.Session.GetString("token");
                client.DefaultRequestHeaders.Add("Authorization", token);
                if (data.Id == 0)
                {
                    var result = client.PostAsync("divisions", byteContent).Result;
                    return Json(result);
                }
                else if (data.Id == id)
                {
                    var result = client.PutAsync("divisions/" + id, byteContent).Result;
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
            var result = client.DeleteAsync("divisions/" + id).Result;
            return Json(result);
        }

        public ActionResult ReportPDF()
        {
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);

            List<Division> divisions = new List<Division>();
            DivisionReport divisonReport = new DivisionReport();
            
            var resTask = client.GetAsync("divisions");
            resTask.Wait();

            var result = resTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<List<Division>>();
                readTask.Wait();
                divisions = readTask.Result;
            }
            else
            {
                divisions = null;
                ModelState.AddModelError(string.Empty, "Server Error try after sometimes.");
            }

            byte[] abytes = divisonReport.PrepareReport(divisions);
            return File(abytes, "application/pdf");
        }

        public async Task<IActionResult> ReportExcel()
        {
            var token = HttpContext.Session.GetString("token");
            client.DefaultRequestHeaders.Add("Authorization", token);

            List<Division> divisions = new List<Division>();
            
            var resTask = client.GetAsync("divisions");
            resTask.Wait();

            var result = resTask.Result;
            if (result.IsSuccessStatusCode)
            {
                var readTask = result.Content.ReadAsAsync<List<Division>>();
                readTask.Wait();
                divisions = readTask.Result;
            }
            else
            {
                divisions = null;
                ModelState.AddModelError(string.Empty, "Server Error try after sometimes.");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("divisions");
                var currentRow = 1;
                var number = 0;

                worksheet.Cell(currentRow, 1).Value = "No";
                worksheet.Cell(currentRow, 2).Value = "Division";
                worksheet.Cell(currentRow, 3).Value = "Department";
                worksheet.Cell(currentRow, 4).Value = "Created Date";
                worksheet.Cell(currentRow, 5).Value = "Updated Date";

                foreach(var form in divisions)
                {
                    currentRow++;
                    number++;
                    worksheet.Cell(currentRow, 1).Value = number;
                    worksheet.Cell(currentRow, 2).Value = form.Name;
                    worksheet.Cell(currentRow, 3).Value = form.Department.Name;
                    worksheet.Cell(currentRow, 4).Value = form.CreateData;
                    worksheet.Cell(currentRow, 5).Value = form.UpdateDate;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var conten = stream.ToArray();
                    return File(conten, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Forms_Data.xlsx");
                }
            }
        }
        
    }
}
