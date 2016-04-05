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

namespace HomeWork.Controllers
{
    public class 客戶資料Controller : Controller
    {
        private 客戶資料Entities1 db = new 客戶資料Entities1();
        客戶資料Repository repo客戶資料 = RepositoryHelper.Get客戶資料Repository();
        客戶聯絡人Repository repo客戶聯絡人 = RepositoryHelper.Get客戶聯絡人Repository();

        // GET: 客戶資料
        public ActionResult Index()
        {
            var listRead = repo客戶資料.All();
            return View(listRead);
        }
                
        [HttpPost]
        public ActionResult Index(string searchWord)
        {
            var list = repo客戶資料.All()
                .Where(c => 
                       (c.客戶名稱.Contains(searchWord) ||
                       c.統一編號.Contains(searchWord) ||
                       c.電話.Contains(searchWord) ||
                       c.傳真.Contains(searchWord) ||
                       c.地址.Contains(searchWord) ||
                       c.Email.Contains(searchWord) ||
                       c.地區.Contains(searchWord)));

            return View(list);
        }

        // GET: 客戶資料/Details/5
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


        public ActionResult 客戶聯絡人BatchUpdate(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ViewData.Model = repo客戶聯絡人.GetData(id.Value);
            ViewBag.客戶Id = id;

            if (ViewData.Model == null)
            {
                return HttpNotFound();
            }
            return PartialView();
        }

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

                //ViewData.Model = repo客戶聯絡人.GetData(ViewBag.客戶Id);
                //ViewBag.客戶Id = ViewBag.客戶Id;

                return PartialView();
            }
            //ViewData.Model = repo客戶聯絡人.GetData(ViewBag.客戶Id);
            return View();
        }

        // GET: 客戶資料/Create
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
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,地區")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                客戶資料.是否已刪除 = false;                
                repo客戶資料.Add(客戶資料);
                repo客戶資料.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            return View(客戶資料);
        }

        // GET: 客戶資料/Edit/5
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
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,地區")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                var db客戶資料 = (客戶資料Entities1)repo客戶資料.UnitOfWork.Context;
                db客戶資料.Entry(客戶資料).State = EntityState.Modified;
                repo客戶資料.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Delete/5
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
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = repo客戶資料.Find(id);
            客戶資料.是否已刪除 = true;
            repo客戶資料.UnitOfWork.Commit();

            return RedirectToAction("Index");
        }

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
