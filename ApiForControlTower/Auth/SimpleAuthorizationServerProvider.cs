using ApiForControlTower.Controllers;
using ApiForControlTower.Models;
using Microsoft.Owin.Security.OAuth;
using System.Security.Claims;
using System.Threading.Tasks;

public class SimpleAuthorizationServerProvider : OAuthAuthorizationServerProvider
{
    public override Task ValidateClientAuthentication(OAuthValidateClientAuthenticationContext context)
    {
        context.Validated();
        return Task.FromResult<object>(null);
    }
    public override async Task GrantResourceOwnerCredentials(OAuthGrantResourceOwnerCredentialsContext context)
    {
        context.OwinContext.Response.Headers.Add("Access-Control-Allow-Origin", new[] { "*" });
        UserController uc = new UserController();
        User user = new User();
        user.UserId = int.Parse(context.UserName);
        user.Password = context.Password;
        User loger = uc.Login(user);
        if (loger == null)
        {
            context.SetError("invalid_grant", "The username or password is incorrect");
            return;
        }
        var identity = new ClaimsIdentity(context.Options.AuthenticationType);
        identity.AddClaim(new Claim("sub", context.UserName));
        identity.AddClaim(new Claim("role", "user"));
        context.Validated(identity);
    }
}