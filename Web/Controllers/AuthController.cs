using BC = BCrypt.Net.BCrypt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using API.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Http;
using System.Net.Http.Headers;
using System.Text;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Session;
using Newtonsoft.Json.Linq;

namespace Web.Controllers
{
    public class AuthController : Controller
    {
        readonly HttpClient client = new HttpClient
        {
            BaseAddress = new Uri("https://localhost:44374/api/")
        };

        [Route("login")]
        public IActionResult Login()
        {
            return View("~/Views/Auth/Login.cshtml");
        }

        [Route("register")]
        public IActionResult Register()
        {
            return View();
        }

        [Route("logout")]
        public IActionResult Logout()
        {
            //var jwtauthmanager = new IJwtAuthManager();
            //var userName = User.Identity.Name;
            //_jwtAuthManager.RemoveRefreshTokenByUserName(userName); // can be more specific to ip, user agent, device name, etc.
            //_logger.LogInformation($"User [{userName}] logged out the system.");

            //HttpContext.Session.Remove("lvl");
            HttpContext.Session.Clear();
            return Redirect("/login");
        }

        [Route("verify")]
        public IActionResult Verify()
        {
            return View();
        }

        [Route("notfound")]
        public IActionResult Notfound()
        {
            return View();
        }

        [Route("validate")]
        public IActionResult Validate(UserVM userVM)
        {
            var json = JsonConvert.SerializeObject(userVM);
            var buffer = System.Text.Encoding.UTF8.GetBytes(json);
            var byteContent = new ByteArrayContent(buffer);
            byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            if (userVM.Username == null)
            { // Login
                HttpResponseMessage result = null;
                if (userVM.VerifyCode != null)
                {
                    result = client.PostAsync("users/code/", byteContent).Result;
                }
                else if (userVM.VerifyCode == null)
                {
                    result = client.PostAsync("users/login/", byteContent).Result;
                }

                if (result.IsSuccessStatusCode)
                {
                    var data = result.Content.ReadAsStringAsync().Result;
                    if (data != null)
                    {
                        HttpContext.Session.SetString("token", "Bearer " + data);
                        var handler = new JwtSecurityTokenHandler();
                        var tokenS = handler.ReadJwtToken(data);
                        var jwtPayloadSer = JsonConvert.SerializeObject(tokenS.Payload.ToDictionary(x => x.Key, x => x.Value));
                        var jwtPayloadDes = JsonConvert.DeserializeObject(jwtPayloadSer).ToString();
                        var account = JsonConvert.DeserializeObject<UserVM>(jwtPayloadSer);

                        if (!account.VerifyCode.Equals(""))
                        {
                            return Json(new { status = true, msg = "VerifyCode" });
                        }
                        else if (account.RoleName != null)
                        {
                            HttpContext.Session.SetString("id", account.Id);
                            HttpContext.Session.SetString("uname", account.Username);
                            HttpContext.Session.SetString("email", account.Email);
                            HttpContext.Session.SetString("lvl", account.RoleName);
                            if (account.RoleName == "Admin")
                            {
                                return Json(new { status = true, msg = "Login Successfully !" });
                                //return View("~/Views/Auth/verify.cshtml");
                            }
                            return Json(new { status = true, msg = "Login Successfully !" });
                        }
                        return Json(new { status = false, msg = "You Don't Have Permissions! Please Contact Administrator" });
                    }
                    return Json(new { status = false, msg = result.Content.ReadAsStringAsync().Result });
                }
                return Json(new { status = false, msg = result.Content.ReadAsStringAsync().Result });
            }
            else if (userVM.Username != null)
            { // Register
                var result = client.PostAsync("users/register/", byteContent).Result;
                if (result.IsSuccessStatusCode)
                {
                    return Json(new { status = true, code = result, msg = "Register Success! " });
                }
                return Json(new { status = false, msg = result.Content.ReadAsStringAsync().Result });
            }
            return Redirect("/login");
        }

        [Route("tes")]
        public IActionResult Tes()
        {
            var stream = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJZCI6ImRiM2VhZmIxLTkyMWUtNDdmYS1hOGFiLTIwNDYxMzkxM2FlMCIsIlVzZXJuYW1lIjoiUmlmcXkiLCJFbWFpbCI6Im11aGFtbWFkcmlmcWkwQGdtYWlsLmNvbSIsIlJvbGVOYW1lIjoiU2FsZXMiLCJWZXJpZnlDb2RlIjoiIiwiZXhwIjoxNTk5NjEyMTYyLCJpc3MiOiJJbnZlbnRvcnlBdXRoZW50aWNhdGlvblNlcnZlciIsImF1ZCI6IkludmVudG9yeXNlcnZpY2VQb3N0bWFudENsaWVudCJ9.UxXzeuzWB8sRuy11Hd1LSzE_oUC7EQ_Ehh4WQMflYXA";
            var handler = new JwtSecurityTokenHandler();
            var tokenS = handler.ReadJwtToken(stream);

            //var jsonToken = handler.ReadToken(stream);
            //var tokenS = handler.ReadToken(stream) as JwtSecurityToken;

            //var user = new UserVM()
            //{
            //    Id = tokenS.Claims.First(claim => claim.Type == "Id").Value,
            //    Username = tokenS.Claims.First(claim => claim.Type == "Username").Value,
            //    Email = tokenS.Claims.First(claim => claim.Type == "Email").Value,
            //    RoleName = tokenS.Claims.First(claim => claim.Type == "RoleName").Value,
            //};

            //var usrVm = new UserVM();
            ////return Json(user);
            //return Json(tokenS.Payload);

            // Re-serialize the Token Headers to just Key and Values
            //var jwtHeader = JsonConvert.SerializeObject(tokenS.Header.Select(h => new { h.Key, h.Value }));
            //jwtOutput = $"{{\r\n\"Header\":\r\n{JToken.Parse(jwtHeader)},";
            // Re-serialize the Token Claims to just Type and Values
            //var jwtPayloadSer = JsonConvert.SerializeObject(tokenS.Payload.Select(c => new { c.Key, c.Value }));
            //jwtOutput = $"\r\n\"Payload\":\r\n{JToken.Parse(jwtPayload)}\r\n}}";

            var jwtPayloadSer = JsonConvert.SerializeObject(tokenS.Payload.ToDictionary(x => x.Key, x => x.Value));
            var jwtPayloadDes = JsonConvert.DeserializeObject(jwtPayloadSer).ToString();
            var account = JsonConvert.DeserializeObject<UserVM>(jwtPayloadSer);

            // Output the whole thing to pretty Json object formatted.
            return Json(new { account.Id, account.Username, account.Email, account.RoleName, account.VerifyCode });
        }

        //[Route("validate")]
        //public IActionResult Validate(UserVM userVM)
        //{
        //    if (userVM.Username == null)
        //    { // Login
        //        var jsonUserVM = JsonConvert.SerializeObject(userVM);
        //        var buffer = System.Text.Encoding.UTF8.GetBytes(jsonUserVM);
        //        var byteContent = new ByteArrayContent(buffer);
        //        byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
        //        var resTask = client.PostAsync("users/login/", byteContent);
        //        resTask.Wait();
        //        var result = resTask.Result;
        //        if (result.IsSuccessStatusCode)
        //        {
        //            var data = result.Content.ReadAsStringAsync().Result;
        //            if (data != null)
        //            {
        //                HttpContext.Session.SetString("token", "Bearer "+data);
        //                var handler = new JwtSecurityTokenHandler();
        //                var tokenS = handler.ReadJwtToken(data);

        //                //var user = new UserVM();
        //                //user.Id = tokenS.Claims.First(claim => claim.Type == "Id").Value;
        //                //user.Username = tokenS.Claims.First(claim => claim.Type == "Username").Value;
        //                //user.Email = tokenS.Claims.First(claim => claim.Type == "Email").Value;
        //                //user.RoleName = tokenS.Claims.First(claim => claim.Type == "RoleName").Value;
        //                //user.VerifyCode = tokenS.Claims.First(claim => claim.Type == "VerifyCode").Value;

        //                var jwtPayloadSer = JsonConvert.SerializeObject(tokenS.Payload.ToDictionary(x => x.Key, x => x.Value));
        //                var jwtPayloadDes = JsonConvert.DeserializeObject(jwtPayloadSer).ToString();
        //                var account = JsonConvert.DeserializeObject<UserVM>(jwtPayloadSer);

        //                //var json = JsonConvert.DeserializeObject(data).ToString();
        //                //var account = JsonConvert.DeserializeObject<UserVM>(json);
        //                //if (BC.Verify(userVM.Password, account.Password) && (account.RoleName == "Admin" || account.RoleName == "Sales"))
        //                if (!account.VerifyCode.Equals(""))
        //                {
        //                    if (userVM.VerifyCode != account.VerifyCode)
        //                    {
        //                        return Json(new { status = true, msg = "Check your Code" });
        //                    }
        //                }
        //                else if (account.RoleName == "Admin" || account.RoleName == "Sales")
        //                {
        //                    HttpContext.Session.SetString("id", account.Id);
        //                    HttpContext.Session.SetString("uname", account.Username);
        //                    HttpContext.Session.SetString("email", account.Email);
        //                    HttpContext.Session.SetString("lvl", account.RoleName);
        //                    if (account.RoleName == "Admin")
        //                    {
        //                        return Json(new { status = true, msg = "Login Successfully !", acc = "Admin" });
        //                    }
        //                    else
        //                    {
        //                        return Json(new { status = true, msg = "Login Successfully !", acc = "Sales" });
        //                    }
        //                }
        //                else
        //                {
        //                    return Json(new { status = false, msg = "Invalid Username or Password!" });
        //                }
        //            }
        //            else
        //            {
        //                return Json(new { status = false, msg = "Username Not Found!" });
        //            }
        //        }
        //        else
        //        {
        //            //return RedirectToAction("Login","Auth");
        //            return Json(new { status = false, msg = "Something Wrong!" });
        //        }
        //    }
        //    else if (userVM.Username != null)
        //    { // Register
        //        var json = JsonConvert.SerializeObject(userVM);
        //        var buffer = System.Text.Encoding.UTF8.GetBytes(json);
        //        var byteContent = new ByteArrayContent(buffer);
        //        byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
        //        var result = client.PostAsync("users/register/", byteContent).Result;
        //        if (result.IsSuccessStatusCode)
        //        {
        //            return Json(new { status = true, code = result, msg = "Register Success! " });
        //        }
        //        else
        //        {
        //            return Json(new { status = false, msg = "Something Wrong!" });
        //        }
        //    }
        //    return Redirect("/login");
        //}

        //[Route("verifCode")]
        //public IActionResult VerifCode(UserVM userVM)
        //{
        //    if (userVM.VerifyCode != null)
        //    {
        //        var jsonUserVM = JsonConvert.SerializeObject(userVM);
        //        var buffer = System.Text.Encoding.UTF8.GetBytes(jsonUserVM);
        //        var byteContent = new ByteArrayContent(buffer);
        //        byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
        //        var result = client.PostAsync("users/code/", byteContent).Result;
        //        if (result.IsSuccessStatusCode)
        //        {
        //            var data = result.Content.ReadAsStringAsync().Result;
        //            if (data != "")
        //            {
        //                var json = JsonConvert.DeserializeObject(data).ToString();
        //                var account = JsonConvert.DeserializeObject<UserVM>(json);
        //                if (account.RoleName == "Admin" || account.RoleName == "Sales")
        //                {
        //                    HttpContext.Session.SetString("id", account.Id);
        //                    HttpContext.Session.SetString("uname", account.Username);
        //                    HttpContext.Session.SetString("email", account.Email);
        //                    HttpContext.Session.SetString("lvl", account.RoleName);
        //                    if (account.RoleName == "Admin")
        //                    {
        //                        return Json(new { status = true, msg = "Login Successfully !", acc = "Admin" });
        //                    }
        //                    else
        //                    {
        //                        return Json(new { status = true, msg = "Login Successfully !", acc = "Sales" });
        //                    }
        //                }
        //                else
        //                {
        //                    return Json(new { status = false, msg = "Invalid Username or Password!" });
        //                }
        //            }
        //            else
        //            {
        //                return Json(new { status = false, msg = "Username Not Found!" });
        //            }
        //            //var data = result.Content.ReadAsStringAsync().Result;
        //            //var json = JsonConvert.DeserializeObject(data).ToString();
        //            //var account = JsonConvert.DeserializeObject<UserVM>(json);
        //            //var dataLogin = new UserVM()
        //            //{
        //            //    Email = account.Email,
        //            //    Password = account.Password
        //            //};
        //            //this.Validate(dataLogin);
        //            //return Json(new { status = true, code = result, msg = "Login Success! " });
        //        }
        //        else
        //        {
        //            return Json(new { status = false, msg = "Your Code is Wrong!" });
        //        }
        //    }
        //    else
        //    {
        //        return Json(new { status = false, msg = "Something Wrong!" });
        //    }
        //}


    }
}
