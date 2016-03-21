using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(HomeWork.Startup))]
namespace HomeWork
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
