Package.describe({
    summary: 'Azure Active Directory OAuth flow',
    version: '1.0.5',
    name: 'justdoinc:azure-active-directory',
    git: 'https://github.com/qipoco/azure-active-directory',
});

Package.onUse(function(api) {
    api.use('oauth2@1.2.0', ['client', 'server']);
    api.use('oauth@1.2.1', ['client', 'server']);
    api.use('http@1.3.0', ['server']);
    api.use(
        ['underscore@1.0.10', 'service-configuration@1.0.11'],
        ['client', 'server'],
    );
    api.use(['random@1.0.10', 'templating@1.3.2'], 'client');
    api.use('accounts-base@1.4.2', ['client', 'server']);

    api.export('AzureAd');

    api.addFiles(['lib/azureAd.js']);
    api.addFiles(
        ['lib/resources.js', 'lib/serverHttp.js', 'resources/graph.js'],
        'server',
    );
    api.addFiles('azure_ad_server.js', 'server');
    api.addFiles('azure_ad_client.js', 'client');

    api.addFiles('azure_ad_accounts.js');
});
