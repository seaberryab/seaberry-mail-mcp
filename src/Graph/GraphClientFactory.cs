using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using SeaberryMailMcp.Auth;

namespace SeaberryMailMcp.Graph;

public sealed class GraphClientFactory
{
    private readonly TokenAcquirer _tokens;

    public GraphClientFactory(TokenAcquirer tokens)
    {
        _tokens = tokens;
    }

    public async Task<GraphServiceClient> CreateAsync(CancellationToken ct)
    {
        // Avoid the type-name collision with Microsoft.Graph.GraphClientFactory.
        var defaultHandlers = Microsoft.Graph.GraphClientFactory.CreateDefaultHandlers();

        // Outermost handler — first to see every outbound request. Blocks all DELETEs
        // before they reach Graph. See ReadOnlyGuardHandler for rationale.
        defaultHandlers.Insert(0, new ReadOnlyGuardHandler());

        var httpClient = Microsoft.Graph.GraphClientFactory.Create(defaultHandlers);

        var provider = new TokenProvider(_tokens);
        var auth = new BaseBearerTokenAuthenticationProvider(provider);

        // The default handlers chain includes RetryHandler (respects Retry-After on 429).
        return new GraphServiceClient(httpClient, auth);
    }

    private sealed class TokenProvider : IAccessTokenProvider
    {
        private readonly TokenAcquirer _tokens;
        public TokenProvider(TokenAcquirer tokens) { _tokens = tokens; }

        public AllowedHostsValidator AllowedHostsValidator { get; } = new();

        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
        {
            var outcome = await _tokens.AcquireAsync(cancellationToken).ConfigureAwait(false);
            return outcome.AccessToken;
        }
    }
}
