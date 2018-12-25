// Request AzureAd credentials for the user
// @param options {optional}
// @param credentialRequestCompleteCallback {Function} Callback function to call on
//   completion. Takes one argument, credentialToken on success, or Error on
//   error.
AzureAd.requestCredential = function (options, credentialRequestCompleteCallback) {
    // support both (options, callback) and (callback).
    if (!credentialRequestCompleteCallback && typeof options === 'function') {
        credentialRequestCompleteCallback = options;
        options = {};
    } else if (!options) {
        options = {};
    }

    var config = AzureAd.getConfiguration(true);
    if (!config) {
        credentialRequestCompleteCallback && credentialRequestCompleteCallback(
            new ServiceConfiguration.ConfigError());
        return;
    }

    var tenant = options.tenant || 'common';
    var scope = (options.scope) ? options.scope.join(' ') : 'user.read';

    var loginStyle = OAuth._loginStyle('azureAd', config, options);
    var credentialToken = Random.secret();

    var queryParams = {
        client_id: config.clientId,
        response_type: 'code',
        redirect_uri: config.redirectUri || OAuth._redirectUri('azureAd', config),
        scope: scope,
        response_mode: 'query',
        state: OAuth._stateParam(loginStyle, credentialToken),
        prompt: options.loginPrompt || 'login',
    };

    if (options.login_hint) {
        queryParams.login_hint = options.login_hint;
    }
    if (options.domain_hint) {
        queryParams.domain_hint = options.domain_hint;
    }

    var queryParamsEncoded = _.map(queryParams, function(val, key) {
        return key + '=' + encodeURIComponent(val);
    });

    var baseUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize?`;
    var loginUrl = baseUrl + queryParamsEncoded.join('&');

    OAuth.launchLogin({
        loginService: "azureAd",
        loginStyle: loginStyle,
        loginUrl: loginUrl,
        credentialRequestCompleteCallback: credentialRequestCompleteCallback,
        credentialToken: credentialToken,
        popupOptions: { height: config.height || 600 }
    });
};
