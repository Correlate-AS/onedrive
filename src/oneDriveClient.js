const { logErrorAndReject, formatDriveResponse } = require('./util.js');
const _ = require('lodash');
const querystring = require('querystring');

const ROOT_URL = 'https://graph.microsoft.com/v1.0'

class OneDriveClient {

    constructor(graphApi, logger) {
        this.graphApi = graphApi;
        this.logger = logger;
    }

    shareTo(fileId, driveId, email) {
        return this.graphApi.request(`${ROOT_URL}/drives/${driveId}/items/${fileId}/invite`, 'POST', {
            requireSignin: true,
            sendInvitation: false,
            roles: ["read"],
            recipients: [{
                email
            }],
            message: "File shared through Correlate"
        })
        .catch(logErrorAndReject('Non-200 while trying to share file', this.logger))
        .then(() => true);
    }

    unshareFrom(fileId, driveId, email) {
        const permissionUrl = `${ROOT_URL}/drives/${driveId}/items/${fileId}/permissions`;
        return this.graphApi.request(permissionUrl)
            .catch(logErrorAndReject('Non-200 while trying to list permissions on file', this.logger))
            .then(data => {
                const permission = data.value.find(d => {
                    return d.invitation && d.invitation.email === email
                })
                if (permission) {
                    return this.graphApi.request(`${permissionUrl}/${permission.id}`, "DELETE")
                    .catch(logErrorAndReject('Non-200 while removing permission', this.logger))
                    .then(() => {});
                }
                this.logger.error("Could not revoke permission from file", { fileId, email });
                return Promise.reject(new Error("Could not revoke permission from file"));
            });
    }

    getAccountId() {
        return this.graphApi.request('https://graph.microsoft.com/v1.0/me/drive/')
        .catch(logErrorAndReject('Non-200 while trying to query user details', this.logger))
        .then(data => data.id);
    }

    searchFiles(query, options = {}) {
        this.logger.info('Searching in OneDrive', { query });
        const qs = querystring.stringify(_.pickBy(options));

        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/root/search(q='${query}')?${qs}`)
            .catch(logErrorAndReject('Non-200 while searching drive', this.logger))
            .then(formatDriveResponse);
    }

    getFilesFrom(parentId, options = {}) {
        parentId = parentId || 'root';
        this.logger.info('Querying OneDrive files', { folder: parentId });
        const qs = querystring.stringify(_.pickBy(options));
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/drive/items/${parentId}/children?${qs}`)
            .catch(logErrorAndReject(`Non-200 while querying folder: ${parentId}`, this.logger))
            .then(formatDriveResponse);
    }

    getFileById(fileId) {
        this.logger.info('Getting OneDrive file', { fileId });
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/drive/items/${fileId}`)
        .catch(logErrorAndReject(`Non-200 while querying file ${fileId}`, this.logger))
        .then(data => {
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

    createFolder(folderName) {
        return this.graphApi.request('https://graph.microsoft.com/v1.0/me/drive/root/children/', 'post', {
            name: folderName,
            folder: { },
            "@microsoft.graph.conflictBehavior": "rename"
        })
            .catch(logErrorAndReject('Non-200 while trying to create folder', this.logger))
            .then(data => {
                return data.name
            });
    }

    createFileAndPopulate(fileName, content, folderId = '') {
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/root:/${folderId}/${fileName}.docx:/content`, 'put', content)
            .catch(logErrorAndReject('Non-200 while trying to create file with content', this.logger))
            .then(data => data.id);
    }
}

module.exports = OneDriveClient;
