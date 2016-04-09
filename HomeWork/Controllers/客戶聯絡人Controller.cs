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
    public class 客戶聯絡人Controller : Controller
    {
        //private 客戶資料Entities1 db = new 客戶資料Entities1();
        客戶聯絡人Repository repo客戶聯絡人 = RepositoryHelper.Get客戶聯絡人Repository();
        客戶資料Repository repo客戶資料 = RepositoryHelper.Get客戶資料Repository();

        // GET: 客戶聯絡人
        public ActionResult Index()
        {
            var listRead = repo客戶聯絡人.All();
            return View(listRead);
        }

        [HttpPost]
        public ActionResult Index(string searchWord)
        {
            var list = repo客戶聯絡人.Search(searchWord);
            return View(list);
        }

        // GET: 客戶聯絡人/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: 客戶聯絡人/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                客戶聯絡人.是否已刪除 = false;
                repo客戶聯絡人.Add(客戶聯絡人);
                repo客戶聯絡人.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // POST: 客戶聯絡人/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                var db客戶聯絡人 = (客戶資料Entities1)repo客戶聯絡人.UnitOfWork.Context;
                db客戶聯絡人.Entry(客戶聯絡人).State = EntityState.Modified;
                repo客戶聯絡人.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: 客戶聯絡人/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // POST: 客戶聯絡人/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id);
            repo客戶聯絡人.Delete(客戶聯絡人);
            repo客戶聯絡人.UnitOfWork.Commit();            
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶聯絡人.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult Export()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();

            // 新增試算表
            ISheet sheet = workbook.CreateSheet("客戶聯絡人");

            #region Header   職稱、姓名、Email、手機、電話、客戶名稱
            sheet.CreateRow(1).CreateCell(0).SetCellValue("職稱");
            sheet.GetRow(1).CreateCell(1).SetCellValue("姓名");
            sheet.GetRow(1).CreateCell(2).SetCellValue("Email");
            sheet.GetRow(1).CreateCell(3).SetCellValue("手機");
            sheet.GetRow(1).CreateCell(4).SetCellValue("電話");
            sheet.GetRow(1).CreateCell(5).SetCellValue("客戶名稱");
            #endregion

            List<客戶聯絡人> list = new List<客戶聯絡人>();
            list = repo客戶聯絡人.All().ToList();

            for (int i = 0; i < list.Count(); i++)
            {
                sheet.CreateRow(i + 2).CreateCell(0).SetCellValue(list[i].職稱);
                sheet.GetRow(i + 2).CreateCell(1).SetCellValue(list[i].姓名);
                sheet.GetRow(i + 2).CreateCell(2).SetCellValue(list[i].Email.ToString());
                sheet.GetRow(i + 2).CreateCell(3).SetCellValue(list[i].手機);
                sheet.GetRow(i + 2).CreateCell(4).SetCellValue(list[i].電話);

                var CustName = repo客戶資料.Find(list[i].客戶Id).客戶名稱.ToString();
                sheet.GetRow(i + 2).CreateCell(5).SetCellValue(CustName);
            }

            workbook.Write(ms);
            workbook = null;
            ms.Close();
            ms.Dispose();

            return File(ms.ToArray(), "application/vnd.ms-excel", "客戶聯絡人.xls");
        }

    }
}
