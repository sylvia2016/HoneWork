using System;
using System.Linq;
using System.Collections.Generic;
	
namespace HomeWork.Models
{   
	public  class 客戶資料Repository : EFRepository<客戶資料>, I客戶資料Repository
	{
        /// <summary>
        /// 取得全部未刪除資料
        /// </summary>
        /// <returns></returns>
        public override IQueryable<客戶資料> All()
        {
            return base.All().Where(c => c.是否已刪除 != true).OrderBy(c => c.客戶名稱);
        }

        /// <summary>
        /// 依據關鍵字搜尋
        /// </summary>
        /// <param name="searchWord">關鍵字</param>
        /// <returns>回傳多筆</returns>
        public IQueryable<客戶資料> Search(string searchWord)
        {
            var queryable = new List<客戶資料>();
            if (!string.IsNullOrEmpty(searchWord))
            {
                queryable = this.All()
                                .Where(c =>
                                      (c.客戶名稱.Contains(searchWord) ||
                                      c.統一編號.Contains(searchWord) ||
                                      c.電話.Contains(searchWord) ||
                                      c.傳真.Contains(searchWord) ||
                                      c.地址.Contains(searchWord) ||
                                      c.Email.Contains(searchWord) ||
                                      c.地區.Contains(searchWord)))
                                 .OrderBy(c => c.客戶名稱)
                                 .ToList();
            }
            
            return queryable.AsQueryable<客戶資料>();
        }

        internal void CheckLogin(string account, string password, out string customId, out string role)
        {
            var 客戶資料 = this.All().Where(c => c.帳號 == account).FirstOrDefault();

            if (password == 客戶資料.密碼)
            {
                if (account == "admin")
                {
                    customId = 客戶資料.Id.ToString();
                    role = "admin";
                }
                else
                {
                    customId = 客戶資料.Id.ToString();
                    role = "member";
                }
            }
            else
            {
                customId = "";
                role = "";
            }
        }

        /// <summary>
        /// 搜尋特定某筆客戶資料
        /// </summary>
        /// <param name="客戶id">客戶ID</param>
        /// <returns></returns>
        public 客戶資料 Find(int 客戶id)
        {
            return this.All().FirstOrDefault(c => c.Id == 客戶id);
        }

        /// <summary>
        /// 刪除客戶資料（只有update"是否已刪除"欄位而已）
        /// </summary>
        /// <param name="entity"></param>
        public override void Delete(客戶資料 entity)
        {
            entity.是否已刪除 = true;
        }
    }

	public  interface I客戶資料Repository : IRepository<客戶資料>
	{

	}
}