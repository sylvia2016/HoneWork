using System;
using System.Linq;
using System.Collections.Generic;
	
namespace HomeWork.Models
{   
	public  class 清單檢視表Repository : EFRepository<清單檢視表>, I清單檢視表Repository
	{

	}

	public  interface I清單檢視表Repository : IRepository<清單檢視表>
	{

	}
}