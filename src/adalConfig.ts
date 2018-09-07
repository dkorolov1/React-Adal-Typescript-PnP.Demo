import { AdalConfig, adalGetToken, AuthenticationContext } from 'react-adal';

// Endpoint URL
export const endpoint = 'https://<tenant name>.sharepoint.com/';
// App Registration ID
const appId = '<App Registration ID>';

export const adalConfig: AdalConfig = {
    cacheLocation: 'localStorage',
    clientId: appId,
    endpoints: {
        api:endpoint
    },
    postLogoutRedirectUri: window.location.origin,
    tenant: '<tenant name>.onmicrosoft.com'
};

class AdalContext {
    private authContext: AuthenticationContext;
    
    constructor() {
        this.authContext = new AuthenticationContext(adalConfig);
    }

    get AuthContext() {
        return this.authContext;
    }

    public GetToken(): Promise<string | null> {
        return adalGetToken(this.authContext, endpoint);
    }

    public LogOut() {
        this.authContext.logOut();
    }
}

const adalContext: AdalContext = new AdalContext();
export default adalContext;