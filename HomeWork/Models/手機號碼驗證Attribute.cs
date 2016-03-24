using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace HomeWork.Models
{
    internal class 手機號碼驗證Attribute : DataTypeAttribute
    {
        public 手機號碼驗證Attribute() :base(DataType.Text)
        {

        }

        public override bool IsValid(object value)
        {
            string str = (string)value;
            bool result = true;

            try
            {
                Regex regex = new Regex(@"\d{4}-\d{6}");
                result = regex.IsMatch(str);
            }
            catch (Exception e)
            {
                throw new Exception("錯誤！", e);
            }

            if (result)
                return true;
            else
                return false;
        }
        
    }
}