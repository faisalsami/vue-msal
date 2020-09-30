'use strict';
import { iMSAL, Options } from './src/types';
import { MSAL } from './src/main';
import { mixin } from "./mixin";

export default class msalPlugin {
    static install(Vue: any, options: Options): void {
        Vue.prototype.$msal = new msalPlugin(options, Vue);
    }
    constructor(options: Options, Vue: any = undefined) {
        const msal = new MSAL(options);
        if (Vue && options.framework && options.framework.globalMixin) {
            Vue.mixin(mixin);
        }
        const exposed: iMSAL = {
            data: msal.data,
            async signIn() { await msal.signIn(); },
            async signOut() { await msal.signOut(); },
            async acquireToken() { await msal.acquireToken(); },
            isAuthenticated() { return msal.isAuthenticated(); }
        };
        return exposed;
    }
}
