using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HomeWork.Models;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;

namespace HomeWork.Controllers
{
    [Authorize(Roles = "admin")]
    public class 客戶銀行資訊Controller : Controller
    {
        //private 客戶資料Entities1 db = new 客戶資料Entities1();
        客戶銀行資訊Repository repo客戶銀行資訊 = RepositoryHelper.Get客戶銀行資訊Repository();
        客戶資料Repository repo客戶資料 = RepositoryHelper.Get客戶資料Repository();

        // GET: 客戶銀行資訊
        public ActionResult Index()
        {
            var listRead = repo客戶銀行資訊.All();
            return View(listRead);
        }

        [HttpPost]
        public ActionResult Index(string searchWord)
        {
            var list = repo客戶銀行資訊.Search(searchWord);
            return View(list);
        }

        // GET: 客戶銀行資訊/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: 客戶銀行資訊/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                客戶銀行資訊.是否已刪除 = false;
                repo客戶銀行資訊.Add(客戶銀行資訊);
                repo客戶銀行資訊.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // POST: 客戶銀行資訊/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                var db客戶銀行資訊 = (客戶資料Entities1)repo客戶銀行資訊.UnitOfWork.Context;
                db客戶銀行資訊.Entry(客戶銀行資訊).State = EntityState.Modified;
                repo客戶銀行資訊.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: 客戶銀行資訊/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // POST: 客戶銀行資訊/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id);
            repo客戶銀行資訊.Delete(客戶銀行資訊);
            repo客戶銀行資訊.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶銀行資訊.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult Export()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();

            // 新增試算表
            ISheet sheet = workbook.CreateSheet("客戶銀行資訊");

            #region Header   銀行名稱、銀行代碼、分行代碼、帳戶名稱、帳戶號碼、客戶名稱
            sheet.CreateRow(1).CreateCell(0).SetCellValue("銀行名稱");
            sheet.GetRow(1).CreateCell(1).SetCellValue("銀行代碼");
            sheet.GetRow(1).CreateCell(2).SetCellValue("分行代碼");
            sheet.GetRow(1).CreateCell(3).SetCellValue("帳戶名稱");
            sheet.GetRow(1).CreateCell(4).SetCellValue("帳戶號碼");
            sheet.GetRow(1).CreateCell(5).SetCellValue("客戶名稱");
            #endregion

            List<客戶銀行資訊> list = new List<客戶銀行資訊>();
            list = repo客戶銀行資訊.All().ToList();

            for (int i = 0; i < list.Count(); i++)
            {
                sheet.CreateRow(i + 2).CreateCell(0).SetCellValue(list[i].銀行名稱);
                sheet.GetRow(i + 2).CreateCell(1).SetCellValue(list[i].銀行代碼);
                sheet.GetRow(i + 2).CreateCell(2).SetCellValue(list[i].分行代碼.ToString());
                sheet.GetRow(i + 2).CreateCell(3).SetCellValue(list[i].帳戶名稱);
                sheet.GetRow(i + 2).CreateCell(4).SetCellValue(list[i].帳戶號碼);

                var CustName = repo客戶資料.Find(list[i].客戶Id).客戶名稱.ToString();
                sheet.GetRow(i + 2).CreateCell(5).SetCellValue(CustName);
            }

            workbook.Write(ms);
            workbook = null;
            ms.Close();
            ms.Dispose();

            return File(ms.ToArray(), "application/vnd.ms-excel", "客戶銀行資訊.xls");
        }

    }
}
