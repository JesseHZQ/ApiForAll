using Microsoft.Owin;
using Microsoft.Owin.Cors;
using Microsoft.Owin.Security.OAuth;
using Owin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace ApiForControlTower
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigAuth(app);
            HttpConfiguration config = new HttpConfiguration();
            WebApiConfig.Register(config);
            app.UseCors(CorsOptions.AllowAll);
            app.UseWebApi(config);
        }
        public void ConfigAuth(IAppBuilder app)
        {
            OAuthAuthorizationServerOptions option = new OAuthAuthorizationServerOptions()
            {
                AllowInsecureHttp = true,
                TokenEndpointPath = new PathString("/token"), //获取 access_token 授权服务请求地址
                AccessTokenExpireTimeSpan = TimeSpan.FromDays(1), //access_token 过期时间
                Provider = new SimpleAuthorizationServerProvider(), //access_token 相关授权服务
                RefreshTokenProvider = new SimpleRefreshTokenProvider() //refresh_token 授权服务
            };
            app.UseOAuthAuthorizationServer(option);
            app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions());
        }
    }
}