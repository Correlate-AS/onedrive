const { addDays } = require('date-fns');
const _ = require('lodash');
const querystring = require('querystring');
const { UPLOAD_CONFLICT_RESOLUTION_MODES } = require('./constants');
const {
    logErrorAndReject,
    formatDriveResponse,
    validateAndDefaultTo,
    getParamValue,
} = require("./util");

const ROOT_URL = 'https://graph.microsoft.com/v1.0'

const SHARING_IDENTIFICATOR = {
    EMAIL: 'email',
    PERMISSION: 'permission',
};

class OneDriveClient {

    constructor(graphApi, logger) {
        this.graphApi = graphApi;
        this.logger = logger;
    }

    shareTo(fileId, driveId, email) {
        return email
            ? this.shareForEmail(fileId, driveId, email)
            : this.shareForAnyone(fileId, driveId)
    }

    shareForEmail(fileId, driveId, email) {
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
    }

    shareForAnyone(fileId, driveId) {
        return this.graphApi.request(`${ROOT_URL}/drives/${driveId}/items/${fileId}/createLink`, 'POST', {
            "type": "view",
            "scope": "anonymous"
          })
        .catch(logErrorAndReject('Non-200 while trying to share file', this.logger));
    }

    /**
     * Unshares file for other user(s) by identificator
     * @async
     *
     * @param {string} fileId File ID
     * @param {string} driveId Drive ID
     * @param {object} identificator Identificator of sharing
     * @param {string} identificator.value Identificator value
     * @param {string} identificator.type Identificator type: 'email', 'permission'
     */
    unshareFrom(fileId, driveId, identificator) {
        const permissionUrl = `${ROOT_URL}/drives/${driveId}/items/${fileId}/permissions`;

        switch (identificator.type) {
            case SHARING_IDENTIFICATOR.EMAIL:
                return this._unshareByEmail(permissionUrl, identificator.value);
            case SHARING_IDENTIFICATOR.PERMISSION:
                return this._unshareByPermission(permissionUrl, identificator.value);
        }
    }

    getAccountId() {
        return this.graphApi.request('https://graph.microsoft.com/v1.0/me/drive/')
        .catch(logErrorAndReject('Non-200 while trying to query user details', this.logger))
        .then(data => data.id);
    }

    getAccountInfo(fields = []) {
        return this.graphApi.request(`${ROOT_URL}/me${fields.length ? `?$select=${fields.join(',')}` : ''}`)
            .catch(logErrorAndReject('Non-200 while trying to query microsoft account details', this.logger));
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
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children?${qs}`)
            .catch(logErrorAndReject(`Non-200 while querying folder: ${parentId}`, this.logger))
            .then(formatDriveResponse);
    }

    getFileById(fileId) {
        this.logger.info('Getting OneDrive file', { fileId });
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`)
        .catch(logErrorAndReject(`Non-200 while querying file ${fileId}`, this.logger))
        .then(data => {
            return {
                name: data.name,
                webUrl: data.webUrl,
                packageType: data.package ? data.package.type : '',
                parentReference: {
                    path: data.parentReference.path,
                },
            }
        });
    }

    getPublicUrl(fileId) {
        return this.getFileById(fileId)
            .then(file => file.webUrl);
    }

    getPreview(fileId) {
        this.logger.info('Getting OneDrive file preview', { fileId });
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/thumbnails`)
            .catch(logErrorAndReject(`Non-200 while querying file ${fileId}`, this.logger))
            .then(data => {
                return  data.value;
            });
    }

    createFolder(folderName, rootFolder = 'root') {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${rootFolder}/children`
        return this.graphApi.request(url, 'post', {
            name: folderName,
            folder: { },
            "@microsoft.graph.conflictBehavior": "rename"
        })
            .catch(logErrorAndReject('Non-200 while trying to create folder', this.logger))
            .then(data => {
                return data.id
            });
    }

    createFileAndPopulate(fileName, content, folderId = '') {
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/items/${folderId}:/${fileName}.docx:/content`, 'put', content)
            .catch(logErrorAndReject('Non-200 while trying to create file with content', this.logger))
            .then(data => data.id);
    }

    listRootFolder() {
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/root/children`)
            .catch(logErrorAndReject('Non-200 while searching drive', this.logger))
            .then(folder => folder.value);
    }

    /**
     * @typedef UploadFileData
     * @property {string} filename
     * @property {ReadableStream} content - stream of a file's content
     */

    /**
     * @param {UploadFileData} fileData
     * @param {string=} parentId - Parent's folder id
     * @param {string=} conflictResolutionMode - Available modes: fail, replace, or rename. {@link https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_put_content?view=odsp-graph-online#conflict-resolution-behavior See docs for more details}
     */
    uploadFile(fileData, parentId , conflictResolutionMode ) {
        conflictResolutionMode = validateAndDefaultTo(conflictResolutionMode, UPLOAD_CONFLICT_RESOLUTION_MODES, UPLOAD_CONFLICT_RESOLUTION_MODES.RENAME);
        parentId = _.defaultTo(parentId, 'root');

        const url = ROOT_URL + `/me/drive/items/${parentId}:/${fileData.filename}:/content?@microsoft.graph.conflictBehavior=${conflictResolutionMode}`;
        return this.graphApi.request(url, 'put', fileData.content)
            .catch(logErrorAndReject('Non-200 while trying to upload file', this.logger));
    }

    /**
     * Returns last cursor of root (recursively - includes state of inner folders)
     */
    getLastCursor() {
        return this.graphApi
          .request(`${ROOT_URL}/me/drive/root/delta?token=latest`)
          .catch(
            logErrorAndReject(
              "Non-200 while trying to get last cursor",
              this.logger
            )
          )
          .then((response) =>
            getParamValue(response["@odata.deltaLink"], "token")
          );
    }

    /**
     * Returns changes performed after the state of provided cursor
     * @param {string} cursor Cursor, which you want to get changes starting from
     */
    getChangesFrom(cursor) {
        return this.graphApi
          .request(`${ROOT_URL}/me/drive/root/delta(token='${cursor}')`)
          .catch(
            logErrorAndReject(
              `Non-200 while trying to get changes from cursor ${cursor}`,
              this.logger
            )
          )
          .then((response) => ({
            ...response,
            cursor: getParamValue(response["@odata.deltaLink"], "token"),
          }));
    }


    /**
     * @typedef CreateSubscriptionPayload
     * @property {string}   [changeType='update']       Indicates the type of change that generated the notification. For OneDrive, this will always be 'updated'
     * @property {string}   notificationUrl             Webhook handler endpoint URL
     * @property {string}   [resource='me/drive/root']  Folder URL, which we subscibe on
     * @property {Date}     [expirationDateTime]        The date and time when the subscription will expire if not updated or renewed (can only be 43200 minutes in the future)
     * @property {string}   [clientState='']            An optional string value that is passed back in the notification message for this subscription.
     */

    /**
     * Creates subscription for folder
     * @param {CreateSubscriptionPayload} payload
     */
    createSubscription(payload) {
        if (!payload.notificationUrl) {
            throw new Error('There was no webhook handler endpoint provided');
        }

        const defaultPaylaod = {
            changeType: 'updated',
            resource: 'me/drive/root',
            expirationDateTime: addDays(new Date(), 30), // 43200 / 60 / 24 = 30 days
            clientState: '',
        };

        const fullPayload = { ...defaultPaylaod, ...payload };

        return this.graphApi.request(`${ROOT_URL}/subscriptions`, 'post', fullPayload)
            .catch(logErrorAndReject('Non-200 while trying to create subscription', this.logger));
    }

    /**
     * @typedef UpdateSubscriptionPayload
     * @property {Date} expirationDateTime Date, which subcription expires on (not more than 43200 hours = 30 days)
     */

    /**
     * Update subscription by id
     * @param {string} subscriptionId ID of subscription, which has to be updated
     * @param {UpdateSubscriptionPayload} payload New data for subscription
     */
     updateSubscription(subscriptionId, payload) {
        return this.graphApi.request(`${ROOT_URL}/subscriptions/${subscriptionId}`, 'patch', payload)
            .catch(logErrorAndReject(`Non-200 while trying to update subscription ${subscriptionId}`, this.logger));
    }

    /**
     * Delete subscription by id
     * @param {string} subscriptionId ID of subscription, which has to be updated
     */
    deleteSubscription(subscriptionId) {
        return this.graphApi.request(`${ROOT_URL}/subscriptions/${subscriptionId}`, 'delete')
            .catch(logErrorAndReject(`Non-200 while trying to delete subscription ${subscriptionId}`, this.logger));
    }

    /**
     * Unshares file by user's email
     * @async
     *
     * @param {string} permissionUrl Permissions endpoint, which can be continued depending on next action
     * @param {string} email User's email
     */
    _unshareByEmail(permissionUrl, email) {
        return this.graphApi.request(permissionUrl)
            .catch(logErrorAndReject('Non-200 while trying to list permissions on file', this.logger))
            .then(data => {
                const permission = data.value.find(d => _.get(d, 'invitation.email') === email);
                if (permission) {
                    return this._unshareByPermission(permissionUrl, permission.id)
                        .then(() => {});
                }
                this.logger.error("Could not revoke permission from file", { fileId, email });
                return Promise.reject(new Error("Could not revoke permission from file"));
            });
    }

    /**
     * Unshares file by permission ID.
     * E.g., public link (for any user) you can unshare only by permissionId
     * @async
     *
     * @param {string} permissionUrl Permissions endpoint, which can be continued depending on next action
     * @param {string} permissionId Permission ID
     */
    _unshareByPermission(permissionUrl, permissionId) {
        return this.graphApi.request(`${permissionUrl}/${permissionId}`, "DELETE")
            .catch(logErrorAndReject('Non-200 while removing permission', this.logger));
    }
    
}

module.exports = OneDriveClient;
