using System;
using System.Linq;
using System.Collections.Generic;
	
namespace HomeWork.Models
{   
	public  class 客戶銀行資訊Repository : EFRepository<客戶銀行資訊>, I客戶銀行資訊Repository
	{
        public override IQueryable<客戶銀行資訊> All()
        {
            return base.All().Where(c => c.是否已刪除 != true);
        }

        public 客戶銀行資訊 Find(int id)
        {
            return this.All().FirstOrDefault(c => c.Id == id);
        }

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