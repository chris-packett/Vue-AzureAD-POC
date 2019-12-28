export const msalConfig = {
    auth: {
        clientId: '<ClientId>',
        authority: 'https://login.microsoftonline.com/<TenantId>'
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: true
    }
};

export const graphConfig = {
    graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me'
};
