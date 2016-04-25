using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HomeWork.Models;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Web.Security;
using System.Linq.Expressions;
using PagedList;

namespace HomeWork.Controllers
{
    [ExecuteTimeLog]
    public class 客戶資料Controller : Controller
    {
        private 客戶資料Entities1 db = new 客戶資料Entities1();
        客戶資料Repository repo客戶資料 = RepositoryHelper.Get客戶資料Repository();
        客戶聯絡人Repository repo客戶聯絡人 = RepositoryHelper.Get客戶聯絡人Repository();

        // GET: 客戶資料
        [Authorize(Roles = "admin")]
        public ActionResult Index(string searchWord, string Sort = "客戶名稱", int Page = 1)
        {
            ViewBag.Sort = Sort;

            //搜尋
            var result = repo客戶資料.Search(searchWord);

            //排序
            var param = Expression.Parameter(typeof(客戶資料), "客戶資料");

            if (!string.IsNullOrEmpty(Sort))
            {
                if (Sort.Contains("DESC"))
                {
                    Sort = Sort.Substring(0, Sort.Length - 5);
                    var expression = Expression.Lambda<Func<客戶資料, object>>(Expression.Property(param, Sort), param);
                    result = result.OrderByDescending(expression);
                }
                else
                {
                    var expression = Expression.Lambda<Func<客戶資料, object>>(Expression.Property(param, Sort), param);
                    result = result.OrderBy(expression);
                }
            }
            
            //分頁
            var data = result.ToPagedList(Page, 3);

            return View(data);
        }

        //[HttpPost]
        //[HandleError(ExceptionType = typeof(ArgumentNullException), View = "Error2")]
        //public ActionResult Index(string searchWord)
        //{
        //    if (string.IsNullOrEmpty(searchWord))
        //    {
        //        throw new ArgumentNullException("您沒有輸入搜尋關鍵字！");
        //    }


        //    var result = repo客戶資料.Search(searchWord);
        //    return View(result);
        //}

        // GET: 客戶資料/Details/5

        [Authorize(Roles = "admin")]
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        [Authorize(Roles = "admin")]
        public ActionResult 客戶聯絡人BatchUpdate(int? 客戶Id)
        {
            if (客戶Id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            ViewData.Model = repo客戶聯絡人.GetDataById(客戶Id.Value);
            ViewBag.客戶Id = 客戶Id;

            if (ViewData.Model == null)
            {
                return HttpNotFound();
            }
            return PartialView();
        }

        [Authorize(Roles = "admin")]
        [HttpPost]
        public ActionResult 客戶聯絡人BatchUpdate(IList<客戶聯絡人Batch> data)
        {
            if (ModelState.IsValid)
            {
                foreach (var item in data)
                {
                    var 客戶聯絡人 = repo客戶聯絡人.Find(item.Id);

                    客戶聯絡人.職稱 = item.職稱;
                    客戶聯絡人.姓名 = item.姓名;
                    客戶聯絡人.Email = item.Email;
                    客戶聯絡人.手機 = item.手機;
                    客戶聯絡人.電話 = item.電話;
                }

                repo客戶聯絡人.UnitOfWork.Commit();
                TempData["Msg"] = "客戶聯絡人更新成功!";                
            }
            //ViewData.Model = repo客戶聯絡人.GetDataById(客戶Id.Value);   repo客戶聯絡人.GetData(ViewBag.客戶Id);
            //ViewBag.客戶Id = ViewBag.客戶Id;

            return RedirectToAction("Index");
        }

        // GET: 客戶資料/Create
        [Authorize(Roles = "admin")]
        public ActionResult Create()
        {
            #region 下拉選單塞值
            List<Area> ls = new List<Area>();
            ls.Add(new Area { AreaId = 0, AreaText = "北部" });
            ls.Add(new Area { AreaId = 1, AreaText = "中部" });
            ls.Add(new Area { AreaId = 2, AreaText = "南部" });
            ls.Add(new Area { AreaId = 3, AreaText = "東部" });
            //IEnumerable accessIDs = ls;
            SelectList sl = new SelectList(ls, "AreaText", "AreaText");
            ViewBag.sl = sl;
            #endregion

            return View();
        }

        // POST: 客戶資料/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [Authorize(Roles = "admin")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,地區,帳號,密碼")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                客戶資料.是否已刪除 = false;
                客戶資料.密碼 = FormsAuthentication.HashPasswordForStoringInConfigFile(客戶資料.密碼, "SHA1");
                repo客戶資料.Add(客戶資料);
                repo客戶資料.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            return View(客戶資料);
        }

        // GET: 客戶資料/Edit/5
        [Authorize]
        public ActionResult Edit(int? id)
        {
            #region 下拉選單塞值
            List<Area> ls = new List<Area>();
            ls.Add(new Area { AreaId = 0, AreaText = "北部" });
            ls.Add(new Area { AreaId = 1, AreaText = "中部" });
            ls.Add(new Area { AreaId = 2, AreaText = "南部" });
            ls.Add(new Area { AreaId = 3, AreaText = "東部" });
            //IEnumerable accessIDs = ls;
            SelectList sl = new SelectList(ls, "AreaText", "AreaText");
            ViewBag.sl = sl;
            #endregion

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: 客戶資料/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int? id, FormCollection form)
        {
            //TODO
            //要想一下怎麼做：如果已經有密碼了，密碼欄位不要空白！現在暫時先以舊密碼再存回去一次
            var 客戶資料 = repo客戶資料.Find(id.Value);
            var oldPassword = 客戶資料.密碼;

            if (TryUpdateModel<客戶資料>(客戶資料, new
                          string[] { "Id", "客戶名稱", "統一編號", "電話", "傳真", "地址", "Email", "地區", "帳號", "密碼"}))
            {
                if (客戶資料.密碼 != null)
                    客戶資料.密碼 = FormsAuthentication.HashPasswordForStoringInConfigFile(客戶資料.密碼, "SHA1");
                else
                    客戶資料.密碼 = oldPassword;  //如果沒有改密碼的話，要把原本的存回去

                repo客戶資料.UnitOfWork.Commit();
                TempData["Msg"] = 客戶資料.客戶名稱 + " 更新成功!";
                return RedirectToAction("Index");
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Delete/5
        [Authorize(Roles = "admin")]
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);

            if (客戶資料 == null)
            {
                return HttpNotFound();
            }

            return View(客戶資料);
        }

        // POST: 客戶資料/Delete/5
        [Authorize(Roles = "admin")]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = repo客戶資料.Find(id);
            repo客戶資料.Delete(客戶資料);
            repo客戶資料.UnitOfWork.Commit();

            return RedirectToAction("Index");
        }

        [Authorize(Roles = "admin")]
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶資料.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult Export()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();

            // 新增試算表
            ISheet sheet = workbook.CreateSheet("客戶資料");
            
            #region Header   客戶名稱、統一編號、電話、傳真、地址、Email、地區
            sheet.CreateRow(1).CreateCell(0).SetCellValue("客戶名稱");
            sheet.GetRow(1).CreateCell(1).SetCellValue("統一編號");
            sheet.GetRow(1).CreateCell(2).SetCellValue("電話");
            sheet.GetRow(1).CreateCell(3).SetCellValue("傳真");
            sheet.GetRow(1).CreateCell(4).SetCellValue("地址");
            sheet.GetRow(1).CreateCell(5).SetCellValue("Email");
            sheet.GetRow(1).CreateCell(6).SetCellValue("地區");
            #endregion

            List<客戶資料> list = new List<客戶資料>();
            list = repo客戶資料.All().ToList();

            for (int i = 0; i < list.Count(); i++)
            {
                sheet.CreateRow(i+2).CreateCell(0).SetCellValue(list[i].客戶名稱);
                sheet.GetRow(i+2).CreateCell(1).SetCellValue(list[i].統一編號);
                sheet.GetRow(i+2).CreateCell(2).SetCellValue(list[i].電話);
                sheet.GetRow(i+2).CreateCell(3).SetCellValue(list[i].傳真);
                sheet.GetRow(i+2).CreateCell(4).SetCellValue(list[i].地址);
                sheet.GetRow(i+2).CreateCell(5).SetCellValue(list[i].Email);
                sheet.GetRow(i+2).CreateCell(6).SetCellValue(list[i].地區);
            }

            workbook.Write(ms);
            workbook = null;
            ms.Close();
            ms.Dispose();

            return File(ms.ToArray(), "application/vnd.ms-excel", "客戶資料.xls");
        }

        class Area
        {
            public int AreaId { get; set; }

            public string AreaText { get; set; }
        }
    }
}
