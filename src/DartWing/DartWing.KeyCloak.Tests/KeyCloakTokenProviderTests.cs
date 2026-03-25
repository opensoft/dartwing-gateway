using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace DartWing.KeyCloak.Tests;

[TestClass]
public class KeyCloakTokenProviderTests
{
    [TestMethod]
    public async Task GetClientToken()
    {
        var configuration = CreateConfiguration();
        RequireSetting(configuration, "KeyCloak:ClientSecret");
        var provider = CreateServiceProvider().GetRequiredService<KeyCloakTokenProvider>();
        var token = await provider.GetClientAccessToken();

        Assert.IsNotNull(token);
        Assert.IsNotEmpty(token.AccessToken);
    }

    [TestMethod]
    public async Task TokenExchange()
    {
        var configuration = CreateConfiguration();
        var userToken = RequireSetting(configuration, "KeyCloakTestAuth:UserToken");
        var exchangeUser = RequireSetting(configuration, "KeyCloakTestAuth:ExchangeUser");
        var provider = CreateServiceProvider().GetRequiredService<KeyCloakTokenProvider>();
        var token = await provider.TokenExchangeUserAccessToken(userToken, "ledger-api");

        Assert.IsNotNull(token);
        Assert.IsNotEmpty(token.AccessToken);

        token = await provider.TokenExchangeUserAccessToken(exchangeUser, "ledger-api");
        Assert.IsNotNull(token);
        Assert.IsNotEmpty(token.AccessToken);
    }

    private static ServiceProvider CreateServiceProvider()
    {
        IServiceCollection services = new ServiceCollection();
        var configurationManager = CreateConfiguration();
        services.AddKeyCloak(configurationManager);
        return services.BuildServiceProvider();
    }

    private static ConfigurationManager CreateConfiguration()
    {
        ConfigurationManager configurationManager = new();
        configurationManager.AddJsonFile("appsettings.json");
        configurationManager.AddJsonFile("appsettings.Development.json", optional: true);
        return configurationManager;
    }

    private static string RequireSetting(IConfiguration configuration, string key)
    {
        var value = configuration[key];

        if (string.IsNullOrWhiteSpace(value))
        {
            Assert.Inconclusive($"Set '{key}' in appsettings.Development.json or environment variables to run this integration test.");
        }

        return value!;
    }
}
