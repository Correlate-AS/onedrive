const _ = require('lodash');
const querystring = require('querystring'); // deprecated for node 14-17, will be stable for node 18
const {
    logErrorAndReject,
    formatDriveResponse,
} = require("./util");
const BaseDriveClient = require('./baseDriveClient');

const ROOT_URL = 'https://graph.microsoft.com/v1.0'
const rootFolderId = 'root';
const SYSTEM_SITES = ['appcatalog'];


class SharepointClient extends BaseDriveClient {

    constructor(graphApi, logger) {
        super(graphApi, logger);
    }

    getAccountInfo(fields = []) {
        return this.graphApi.request(`${ROOT_URL}/me${fields.length ? `?$select=${fields.join(',')}` : ''}`)
            .catch(logErrorAndReject('Non-200 while trying to query microsoft account details', this.logger));
    }

    async getAccountId() {
        return this.getAccountInfo().then(data => data.id);
    }

    getFilesFrom(siteId, parentId, options = {}) {
        parentId = parentId || rootFolderId;
        siteId = siteId || rootFolderId;

        this.logger.info('Querying Sharepoint files', { site: siteId, folder: parentId });
        const qs = querystring.stringify(_.pickBy(options));
        return this.graphApi.request(`${ROOT_URL}/sites/${siteId}/drive/items/${parentId}/children?${qs}`)
            .catch(logErrorAndReject(`Non-200 while querying folder: ${parentId}`, this.logger))
            .then(formatDriveResponse);
    }

    getPreview(fileId, siteId) {
        siteId = siteId || rootFolderId;

        this.logger.info('Getting Sharepoint file preview', { siteId, fileId });
        return super.getPreview(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}/thumbnails`);
    }

    getSites() {
        this.logger.info('Getting Sharepoint sites');

        return this.graphApi.request(`${ROOT_URL}/sites?search=`)
            .catch(logErrorAndReject('Non-200 while getting sites', this.logger))
            .then(response => ({ // filter system sites
                ...response,
                value: response.value.filter(v => !SYSTEM_SITES.includes(v.name)),
            }))
            .then(formatDriveResponse);
    }

    getFileById(fileId, siteId, options) {
        siteId = siteId || rootFolderId;

        this.logger.info(`Getting Sharepoint file`, { siteId, fileId });
        return super.getFileById(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}`, options);
    }

    getPublicUrl(fileId, siteId) {
        return this.getFileById(fileId, siteId)
            .then(file => file.webUrl);
    }

    shareForEmail(fileId, siteId, email) {
        return super.shareForEmail(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}/invite`, email);
    }

    unshareFrom(fileId, siteId, permissionId) {
        return super.unshareFrom(`${ROOT_URL}/sites/${siteId}/drive/items/${fileId}/permissions/${permissionId}`);
    }
}

module.exports = SharepointClient;
