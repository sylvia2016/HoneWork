using HomeWork.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace HomeWork.Controllers
{
    public class AccountController : Controller
    {
        string gUserData = "";
        string GCustomId = "";

        [AllowAnonymous]
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public ActionResult Login(LoginViewModel data)
        {
            string strPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(data.Password, "SHA1");

            if (CheckLogin(data.Account, strPassword))
            {
                // 將管理者登入的 Cookie 設定成 Session Cookie
                bool isPersistent = false;

                FormsAuthenticationTicket ticket = new FormsAuthenticationTicket(1,
                  data.Account,
                  DateTime.Now,
                  DateTime.Now.AddMinutes(30),
                  isPersistent,
                  gUserData,
                  FormsAuthentication.FormsCookiePath);

                string encTicket = FormsAuthentication.Encrypt(ticket);

                // Create the cookie.
                Response.Cookies.Add(new HttpCookie(FormsAuthentication.FormsCookieName, encTicket));

                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View("Login", "Account");
                //Page.ClientScript.RegisterStartupScript(Page.GetType(), "LoginFail", "alert('登入失敗，請重新登入!');", true);
            }
        }

        private bool CheckLogin(string account, string password)
        {
            string customId;
            string role;
            客戶資料Repository repo客戶資料 = RepositoryHelper.Get客戶資料Repository();
            repo客戶資料.CheckLogin(account, password, out customId, out role);

            gUserData = role;
            GCustomId = customId;

            if (string.IsNullOrEmpty(customId))
                return false;
            else
                return true;
        }

        [AllowAnonymous]
        public ActionResult Register()
        {
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public ActionResult Register(RegisterViewModel register)
        {
            if (ModelState.IsValid)
            {
                //TODO

                RedirectToAction("Index", "Home");
            }
            return View();
        }

        [AllowAnonymous]
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Index", "Home");
        }

        public ActionResult EditProfile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult EditProfile(EditProfileViewModel data)
        {
            if (ModelState.IsValid)
            {
                //TODO

                RedirectToAction("Index", "Home");
            }
            return View();
        }
    }
}