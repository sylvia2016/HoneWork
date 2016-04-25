using HomeWork.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HomeWork.Controllers
{
    [Authorize(Roles = "admin")]
    public class 清單檢視表Controller : Controller
    {
        客戶資料Entities1 db = new 客戶資料Entities1();

        // GET: 清單檢視表
        public ActionResult Index()
        {
            var data = db.清單檢視表.ToList();
            return View(data);
        }
    }
}