import conf from "msal/lib-commonjs/Configuration";
import msal from "msal";

export type CacheOptions = conf.CacheOptions;
export type AuthError = msal.AuthError;
export type AuthResponse = msal.AuthResponse;
export type SystemOptions = conf.SystemOptions;
export type Account = msal.Account;

export type DataObject = {
    isAuthenticated: boolean,
    accessToken: string,
    idToken: string,
    user: User,
    custom: object,
    account?: msal.Account
}

export type FrameworkOptions = {
    globalMixin?: boolean
}


export type Options = {
    auth: Auth,
    loginRequest: Request,
    tokenRequest: Request,
    cache?: CacheOptions,
    system?: SystemOptions,
    framework?: FrameworkOptions
}

export type Request = {
    scopes?: string[]
    account? : msal.Account
}

// Config object to be passed to Msal on creation.
// For a full list of msal.js configuration parameters, 
// visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
export type Auth = {
    clientId: string,
    authority: string,
    redirectUri: string,
    autoRefreshToken?: boolean,
    onAuthentication: (ctx: object, error: AuthError, response: AuthResponse) => any,
    onToken: (ctx: object, error: AuthError | null, response: AuthResponse | null) => any,
    beforeSignOut: (ctx: object) => any
}

export interface iMSAL {
    data: DataObject,
    signIn: () => Promise<any> | void,
    signOut: () => Promise<any> | void,
    acquireToken: () => Promise<any> | void,
    isAuthenticated: () => boolean
}

export type User = {
    name: string,
    userName: string
}
