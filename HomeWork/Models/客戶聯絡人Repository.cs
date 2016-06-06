using System;
using System.Linq;
using System.Collections.Generic;
using System.Web.Mvc;

namespace HomeWork.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        /// <summary>
        /// 取得全部未刪除資料
        /// </summary>
        /// <returns></returns>
        public override IQueryable<客戶聯絡人> All()
        {
            var result = base.All().Where(c => c.是否已刪除 != true).OrderBy(c => c.職稱);
            return result;
        }

        /// <summary>
        /// 依據關鍵字搜尋
        /// </summary>
        /// <param name="searchWord">關鍵字</param>
        /// <returns>回傳多筆</returns>
        public IQueryable<客戶聯絡人> Search(string searchWord)
        {
            var queryable = new List<客戶聯絡人>();

            if (!string.IsNullOrEmpty(searchWord))
            {
                queryable = this.All()
                                .Where(c =>
                                      (c.職稱.Contains(searchWord) ||
                                       c.姓名.Contains(searchWord) ||
                                       c.Email.Contains(searchWord) ||
                                       c.手機.Contains(searchWord) ||
                                       c.電話.Contains(searchWord) ||
                                       c.客戶資料.客戶名稱.Contains(searchWord)))                                
                                .ToList();
            }
            else
            {
                queryable = this.All().ToList();
            }

            return queryable.AsQueryable<客戶聯絡人>();
        }

        internal IQueryable<客戶聯絡人> Filter(string ddlData)
        {
            var queryable = new List<客戶聯絡人>();

            if (!string.IsNullOrEmpty(ddlData))
            {
                queryable = this.All()
                                .Where(c => c.職稱.Contains(ddlData))
                                .ToList();
            }
            else
            {
                queryable = this.All().ToList();
            }

            return queryable.AsQueryable<客戶聯絡人>();
        }

        internal IQueryable GetDDLdata()
        {
            var result = this.All().Select(c => new { c.職稱 }).Distinct();
            return result;
        }

        /// <summary>
        /// 依據客戶聯絡人ID搜尋
        /// </summary>
        /// <param name="id">客戶聯絡人ID</param>
        /// <returns></returns>
        public 客戶聯絡人 Find(int 聯絡人id)
        {
            return this.All().FirstOrDefault(c => c.Id == 聯絡人id);
        }

        /// <summary>
        /// 依據客戶ID取得該客戶的所有聯絡人資料
        /// </summary>
        /// <param name="客戶Id">客戶Id</param>
        /// <returns>多筆</returns>
        public IQueryable<客戶聯絡人> GetDataById(int? 客戶Id)
        {
            if (客戶Id.HasValue)
                return this.All().Where(c => c.客戶Id == 客戶Id);
            else
                return null;
        }

        /// <summary>
        /// 刪除客戶聯絡人（只有update"是否已刪除"欄位而已）
        /// </summary>
        /// <param name="entity"></param>
        public override void Delete(客戶聯絡人 entity)
        {
            //base.Delete(entity);
            entity.是否已刪除 = true;
        }
    }

	public  interface I客戶聯絡人Repository : IRepository<客戶聯絡人>
	{

	}
}