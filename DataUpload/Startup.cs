using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(DataUpload.Startup))]
namespace DataUpload
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
