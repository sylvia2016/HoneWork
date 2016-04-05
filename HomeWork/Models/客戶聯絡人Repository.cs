using System;
using System.Linq;
using System.Collections.Generic;
	
namespace HomeWork.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        public override IQueryable<客戶聯絡人> All()
        {
            return base.All().Where(c => c.是否已刪除 != true);
        }

        public 客戶聯絡人 Find(int id)
        {
            return this.All().FirstOrDefault(c => c.Id == id);
        }

        public IQueryable<客戶聯絡人> GetData(int id)
        {
            return this.All().Where(c => c.客戶Id == id);
        }

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