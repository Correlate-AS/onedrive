const { addDays } = require('date-fns');
const _ = require('lodash');
const querystring = require('querystring'); // deprecated for node 14-17, will be stable for node 18 
const {
    logErrorAndReject,
    formatDriveResponse,
} = require("./util");

const ROOT_URL = 'https://graph.microsoft.com/v1.0'
const rootFolderId = 'root';

// TODO: create base class for sharepoint and onedrive

class SharepointClient {

    constructor(graphApi, logger) {
        this.graphApi = graphApi;
        this.logger = logger;
    }

    getDriveInfo(fields = []) {
        const qs = fields.length
            ? querystring.stringify({ '$select': fields.join(',') })
            : '';

        return this.graphApi.request(`${ROOT_URL}/sites/root/drive?${qs}`)
            .catch(logErrorAndReject('Non-200 while trying to query user details', this.logger));
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

        this.logger.info('Querying Sharepoint files', { folder: parentId });
        const qs = querystring.stringify(_.pickBy(options));
        return this.graphApi.request(`${ROOT_URL}/sites/${siteId}/drive/items/${parentId}/children?${qs}`)
            .catch(logErrorAndReject(`Non-200 while querying folder: ${parentId}`, this.logger))
            .then(formatDriveResponse);
    }

    getPreview(fileId) {
        this.logger.info('Getting Sharepoint file preview', { fileId });
        return this.graphApi.request(`${ROOT_URL}/sites/root/drive/items/${fileId}/thumbnails`)
            .catch(logErrorAndReject(`Non-200 while querying file ${fileId}`, this.logger))
            .then(data => {
                return  data.value;
            });
    }

    searchFiles(query, options = {}) {
        this.logger.info('Searching in Sharepoint', { query });
        const qs = querystring.stringify(_.pickBy(options));

        return this.graphApi.request(`${ROOT_URL}/sites/root/drive/root/search(q='${query}')?${qs}`)
            .catch(logErrorAndReject('Non-200 while searching drive', this.logger))
            .then(formatDriveResponse);
    }

    getSites() {
        this.logger.info('Getting Sharepoint sites');

        return this.graphApi.request(`${ROOT_URL}/sites?search=`)
            .catch(logErrorAndReject('Non-200 while getting sites', this.logger))
            .then(formatDriveResponse);
    }
}

module.exports = SharepointClient;
