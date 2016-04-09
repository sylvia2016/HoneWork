using System;
using System.Linq;
using System.Collections.Generic;
	
namespace HomeWork.Models
{   
	public  class 客戶銀行資訊Repository : EFRepository<客戶銀行資訊>, I客戶銀行資訊Repository
	{
        /// <summary>
        /// 取得全部未刪除資料
        /// </summary>
        /// <returns></returns>
        public override IQueryable<客戶銀行資訊> All()
        {
            return base.All().Where(c => c.是否已刪除 != true);
        }

        /// <summary>
        /// 依據關鍵字搜尋
        /// </summary>
        /// <param name="searchWord">關鍵字</param>
        /// <returns>回傳多筆</returns>
        public IQueryable<客戶銀行資訊> Search(string searchWord)
        {
            var queryable = new List<客戶銀行資訊>();
            if (!string.IsNullOrEmpty(searchWord))
            {
                queryable = this.All()
                                .Where(c =>
                                       c.銀行名稱.Contains(searchWord) ||
                                       c.帳戶名稱.Contains(searchWord) ||
                                       c.銀行代碼.ToString().Contains(searchWord) ||
                                       c.分行代碼.ToString().Contains(searchWord) ||
                                       c.帳戶號碼.Contains(searchWord)).ToList();
            }

            return queryable.AsQueryable<客戶銀行資訊>();
        }

        /// <summary>
        /// 搜尋特定某筆客戶資料
        /// </summary>
        /// <param name="銀行id">銀行ID</param>
        /// <returns></returns>
        public 客戶銀行資訊 Find(int 銀行id)
        {
            return this.All().FirstOrDefault(c => c.Id == 銀行id);
        }

        /// <summary>
        /// 刪除客戶銀行資訊（只有update"是否已刪除"欄位而已）
        /// </summary>
        /// <param name="entity"></param>
        public override void Delete(客戶銀行資訊 entity)
        {
            //base.Delete(entity);
            entity.是否已刪除 = true;
        }

    }

    public  interface I客戶銀行資訊Repository : IRepository<客戶銀行資訊>
	{

	}
}