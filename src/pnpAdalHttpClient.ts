import { adalFetch, AuthenticationContext } from 'react-adal';

import { HttpClientImpl } from '@pnp/common';

export class PnpAdalHttpClient implements HttpClientImpl {
    
    private endpoint: string;
    private authContext: AuthenticationContext;

    constructor (endpoint: string, authContext: AuthenticationContext) {
        this.endpoint = endpoint;
        this.authContext = authContext;
    }

    public fetch(url: string, options: any): Promise<Response> {
        
        let resultOptions = options;
        if(options.headers && typeof options.headers[Symbol.iterator] === 'function') {
            const headers: any = {};
             
            const headerValues = Array.from(options.headers);
            headerValues.forEach(value => {
                const headerName = value[0];
                const headerValue = value[1];
                headers[headerName] = headerValue;
            });

            resultOptions = { ...options, headers };
        }

        return adalFetch(this.authContext, this.endpoint, fetch, url, resultOptions);
    }
}