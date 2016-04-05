using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace HomeWork.Models
{
    public class 客戶聯絡人Batch
    {
        public int Id { get; set; }
        public int 客戶Id { get; set; }
        public string 職稱 { get; set; }
        public string 姓名 { get; set; }
        public string Email { get; set; }
        public string 手機 { get; set; }
        public string 電話 { get; set; }
    }
}