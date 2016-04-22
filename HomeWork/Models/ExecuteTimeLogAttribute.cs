using System;
using System.Diagnostics;
using System.Web.Mvc;

namespace HomeWork.Models
{
    internal class ExecuteTimeLogAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.ActionStartTime = DateTime.Now;
            base.OnActionExecuting(filterContext);
        }

        public override void OnActionExecuted(ActionExecutedContext filterContext)
        {
            filterContext.Controller.ViewBag.ActionEndTime = DateTime.Now;

            TimeSpan ActionTimeSpan = (filterContext.Controller.ViewBag.ActionEndTime
                                    - filterContext.Controller.ViewBag.ActionStartTime);

            //filterContext.Controller.ViewBag.ActionTimeSpan = ActionTimeSpan;
            Debug.WriteLine("Action執行時間：共花了 " + ActionTimeSpan.TotalSeconds + " 秒");

            base.OnActionExecuted(filterContext);
        }

        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.ResultStartTime = DateTime.Now;
            base.OnResultExecuting(filterContext);
        }

        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {
            filterContext.Controller.ViewBag.ResultEndTime = DateTime.Now;

            TimeSpan ResultTimeSpan = filterContext.Controller.ViewBag.ResultEndTime
                                    - filterContext.Controller.ViewBag.ResultStartTime;

            //filterContext.Controller.ViewBag.ResultTimeSpan = ResultTimeSpan;
            Debug.WriteLine("Result執行時間：共花了 " + ResultTimeSpan.TotalSeconds + " 秒");
            
            base.OnResultExecuted(filterContext);
        }
    }
}