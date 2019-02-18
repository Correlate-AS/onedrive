const axios = require('axios');
const { logErrorAndReject } = require('./util.js');
const _ = require('lodash');

const URL = 'https://graph.microsoft.com/v1.0'

class OneDriveClient {

    constructor(accessToken, refreshToken, performRefresh, logger) {
        this.accessToken = accessToken;
        this.refreshToken = refreshToken;
        this.performRefresh = performRefresh;
        this.logger = logger;
    }

    request(url, method = 'get', data) {
        const doIt = () => axios({
            method,
            url,
            data,
            headers: { 'Authorization': `Bearer ${this.accessToken}` }
        });
        return doIt()
        .catch(e => {
            const message = _.get(e, 'response.data.error.message');
            if (message && message.match(/code: 80049228/i)) {
                if (this.performRefresh) {
                    this.logger.info('Access token expired, trying to refresh');
                    return this.performRefresh(this.refreshToken)
                        .then((tokens) => {
                            this.accessToken = tokens.accessToken;
                            this.refreshToken = tokens.refreshToken;
                        })
                        .catch(e => {
                            this.logger.error("Could not use the refresh token", { response: e.response.data });
                            return Promise.reject(new Error("Could not use refresh token to get a new access token"));
                        })
                        .then(doIt);
                }
                this.logger.error('Access token expired, and no refresh mechanism is set');
                return Promise.reject(new Error('Access token has expired'));
            }
            return Promise.reject(e);
        });
    }

    shareTo(fileId, driveId, email) {
        return this.request(`${URL}/drives/${driveId}/items/${fileId}/invite`, 'POST', {
            requireSignin: true,
            sendInvitation: false,
            roles: ["read"],
            recipients: [{
                email
            }],
            message: "File shared through Correlate"
        })
        .catch(logErrorAndReject('Non-200 while trying to share file', this.logger))
        .then(() => {
            return true;
        })
    }

    unshareFrom(fileId, driveId, email) {
        const permissionUrl = `${URL}/drives/${driveId}/items/${fileId}/permissions`;
        return this.request(permissionUrl)
            .catch(logErrorAndReject('Non-200 while trying to list permissions on file', this.logger))
            .then(({ data }) => {
                const permission = data.value.find(d => {
                    return d.invitation && d.invitation.email === email
                })
                if (permission) {
                    return this.request(`${permissionUrl}/${permission.id}`, "DELETE")
                    .catch(logErrorAndReject('Non-200 while removing permission', this.logger))
                    .then(() => {});
                }
                this.logger.error("Could not revoke permission from file", { fileId, email });
                return Promise.reject(new Error("Could not revoke permission from file"));
            });
    }

    getAccountId() {
        return this.request('https://graph.microsoft.com/v1.0/me/drive/')
        .catch(logErrorAndReject('Non-200 while trying to query user details', this.logger))
        .then(({ data }) => {
            return data.id;
        })
    }

    getFilesFrom(parentId) {
        parentId = parentId || 'root';
        this.logger.info('Querying OneDrive files', { folder: parentId });
        return this.request(`https://graph.microsoft.com/v1.0/drive/items/${parentId}/children`)
        .catch(logErrorAndReject(`Non-200 while querying folder: ${parentId}`, this.logger))
        .then(({ data }) => {
            const { value } = data;
            return {
                cursor: null,
                files: value.map(file => ({
                    id: file.id,
                    isFolder: !!file.folder,
                    name: file.name
                }))
            };
        })
    }

    getFileById(fileId) {
        this.logger.info('Getting OneDrive file', { fileId });
        return this.request(`https://graph.microsoft.com/v1.0/drive/items/${fileId}`)
        .catch(logErrorAndReject(`Non-200 while querying file ${fileId}`, this.logger))
        .then(response => {
            const { data } = response;
            return {
                name: data.name,
                webUrl: data.webUrl
            }
        });
    }

    getPublicUrl(fileId) {
        return this.getFileById(fileId)
            .then(file => file.webUrl);
    }

}

module.exports = OneDriveClient;
