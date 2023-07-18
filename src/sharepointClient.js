const {
    logErrorAndReject,
    formatDriveResponse,
} = require("./util");
const BaseDriveClient = require('./baseDriveClient');

const ROOT_URL = 'https://graph.microsoft.com/v1.0'
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

    getFilesFrom(parentId, options = {}, siteId) {
        parentId = this._validateContainer(parentId);
        siteId = this._validateContainer(siteId);

        this.logger.info('Querying Sharepoint files', { site: siteId, folder: parentId });
        return super.getFilesFrom(`${this.ROOT_URL}/sites/${siteId}/drive/items/${parentId}/children`, options);
    }

    getPreview(fileId, siteId) {
        siteId = this._validateContainer(siteId);

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
        siteId = this._validateContainer(siteId);

        this.logger.info(`Getting Sharepoint file`, { siteId, fileId });
        return super.getFileById(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}`, options);
    }

    getFileContent(fileId, siteId, options, axiosOptions) {
        this.logger.info(`Getting OneDrive file content`, { fileId });
        return super.getFileContent(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}/content`, options, axiosOptions);
    }

    getPublicUrl(fileId, siteId) {
        this.logger.info(`Getting Sharepoint sharing link`, { siteId, fileId });

        return this.createSharingLink(fileId, siteId)
            .then(permission => permission.link.webUrl);
    }

    /**
     * Creates sharing link of specified type if it does not already exist
     * or returns sharing link if link of such a type already exists
     * @param {string} fileId Sharepoint drive item ID
     * @param {string} siteId Sharepoint site ID
     * @returns {Promise<Permission>} https://learn.microsoft.com/en-us/graph/api/resources/permission?view=graph-rest-1.0#properties
     * @async
     */
    createSharingLink(fileId, siteId) {
        siteId = this._validateContainer(siteId);
        this.logger.info(`Creating Sharepoint sharing link (or getting if it already exists)`, { siteId, fileId });

        return super.createSharingLink(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}/createLink`);
    }

    /**
     * Deletes file's permission
     * @param {string} permissionId Permission ID of Sharepoint drive item (item can have multiple permissions)
     * @param {string} fileId Sharepoint drive item ID
     * @param {string} siteId Sharepoint site ID
     * @returns {Promise<undefined>}
     * @async
     */
    deletePermission(permissionId, fileId, siteId) {
        siteId = this._validateContainer(siteId);
        this.logger.info(`Removing Sharepoint file permission`, { siteId, fileId, permissionId });

        return super.deletePermissions(`${this.ROOT_URL}/sites/${siteId}/drive/items/${fileId}/permissions/${permissionId}`);
    }
}

module.exports = SharepointClient;
