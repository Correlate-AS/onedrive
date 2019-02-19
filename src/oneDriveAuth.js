const querystring = require('querystring');
const axios = require('axios');
const { logErrorAndReject } = require('./util.js');
const { DEFAULT_SCOPES } = require('./util.js');

class OneDriveAuth {

    constructor(clientId, clientSecret, callbackUrl, logger, scopes = DEFAULT_SCOPES) {
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.scopes = scopes;
        this.callbackUrl = callbackUrl;
        this.logger = logger;
    }

    generateAuthUrl() {
        const params = querystring.stringify({
            scope: this.scopes.join(' '),
            client_id: this.clientId,
            redirect_uri: this.callbackUrl,
            response_type: 'code',
        });

        return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params}`;
    }

    tokensFromCode(authCode) {
        const qs = querystring.stringify({
            grant_type: 'authorization_code',
            client_id: this.clientId,
            client_secret: this.clientSecret,
            scope: this.scopes.join(" "),
            code: authCode
        });
        return axios({
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            data: qs,
            method: 'POST'
        })
        .catch(logErrorAndReject('Non-200 while exchanging auth code for access token', this.logger))
        .then(({ data }) => {
            return {
                accessToken: data.access_token,
                refreshToken: data.refresh_token
            }
        });
    }

    refresh(refreshToken) {
        const data = querystring.stringify({
            grant_type: 'refresh_token',
            client_id: this.clientId,
            client_secret: this.clientSecret,
            scope: this.scopes.join(" "),
            refresh_token: refreshToken
        });
        return axios(`https://login.microsoftonline.com/common/oauth2/v2.0/token`, {
            method: 'post',
            data,
            headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=utf-8' },
        }).then(response => {
            this.logger.info('Successfully used refresh token');
            const { data } = response;
            const tokens = { accessToken: data.access_token, refreshToken: data.refresh_token };
            return tokens;
        });
    }

}

module.exports = OneDriveAuth;
