const axios = require('axios');
const { get } = require('lodash');
const { logErrorAndReject } = require('./util.js');

class GraphAPI {

    constructor(accessToken, refreshToken, graphAuth, onRefresh, logger) {
        this.accessToken = accessToken;
        this.refreshToken = refreshToken;
        this.graphAuth = graphAuth;
        this.onRefresh = onRefresh;
        this.logger = logger;
    }

    request(url, method = 'get', data) {
        const doRequest = () => axios({
            method,
            url,
            data,
            headers: { 'Authorization': `Bearer ${this.accessToken}` }
        });
        return doRequest()
        .catch(e => {
            const message = get(e, 'response.data.error.message');
            if (message && message.match(/code: 80049228/i)) {
                if (this.graphAuth) {
                    this.logger.info('Access token expired, trying to refresh');
                    return this.graphAuth.refresh(this.refreshToken)
                    .catch(e => {
                        this.logger.error("Could not use the refresh token", { response: e.response.data });
                        return Promise.reject(new Error("Could not use refresh token to get a new access token"));
                    })
                    .then(({ accessToken, refreshToken }) => {
                        this.logger.info('Successful refresh');
                        this.accessToken = accessToken;
                        this.refreshToken = refreshToken;
                        if (this.onRefresh) {
                            return Promise.resolve(this.onRefresh(accessToken, refreshToken))
                        }
                    })
                    .then(doRequest);
                }
                this.logger.error("The access token has expired, and no auth client is set to renew it");
            }
            return Promise.reject(e);
        })
        .catch(logErrorAndReject('Received a non-200 while executing a Graph API call', this.logger))
        .then(r => r.data)
    }
}

module.exports = GraphAPI;
