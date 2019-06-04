const querystring = require('querystring');
const axios = require('axios');
const { logErrorAndReject } = require('./util.js');

const OAUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0';

class GraphAuth {

    constructor(clientId, clientSecret, callbackUrl, scopes, logger) {
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.callbackUrl = callbackUrl;
        this.scopes = scopes;
        this.logger = logger;
    }

    generateAuthUrl(state) {
        const params = {
            scope: this.scopes.join(' '),
            client_id: this.clientId,
            redirect_uri: this.callbackUrl,
            prompt: 'consent',
            response_type: 'code',
        };
        if (state) {
            params.state = state;
        }
        this.logger.info(params, "Generating OAuth url");

        return `${OAUTH_URL}/authorize?${querystring.stringify(params)}`;
    }

    tokensFromCode(authCode) {
        const qs = querystring.stringify({
            redirect_uri: this.callbackUrl,
            grant_type: 'authorization_code',
            client_id: this.clientId,
            client_secret: this.clientSecret,
            scope: this.scopes.join(" "),
            code: authCode
        });
        return axios({
            url: `${OAUTH_URL}/token`,
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
        return axios(`${OAUTH_URL}/token`, {
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

module.exports = GraphAuth;
