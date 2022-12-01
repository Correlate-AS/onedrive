const { addDays } = require('date-fns');
const _ = require('lodash');
const querystring = require('querystring'); // deprecated for node 14-17, will be stable for node 18
const { UPLOAD_CONFLICT_RESOLUTION_MODES } = require('./constants');
const {
    logErrorAndReject,
    formatDriveResponse,
    validateAndDefaultTo,
    getParamValue,
} = require("./util");
const BaseDriveClient = require('./baseDriveClient');


const ROOT_URL = 'https://graph.microsoft.com/v1.0';
const rootFolderId = 'root';

class OneDriveClient extends BaseDriveClient {

    constructor(graphApi, logger) {
        super(graphApi, logger);
    }

    shareTo(fileId, driveId, email) {
        return email
            ? this.shareForEmail(fileId, driveId, email)
            : this.shareForAnyone(fileId, driveId)
    }

    /**
     * Shares drive item for email
     * @param {string} fileId Drive item ID
     * @param {string} driveId Drive ID, which contains item
     * @param {string} email Email, which share to
     * @returns {Promise}
     * @async
     */
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
     * Unshares drive item for email
     * @param {string} fileId Drive item ID
     * @param {string} driveId Drive ID, which contains item
     * @param {string} email Permission ID, which allows sharing
     * @returns {Promise}
     * @async
     */
     unshareFrom(fileId, driveId, email) {
        const permissionUrl = `${ROOT_URL}/drives/${driveId}/items/${fileId}/permissions`;
        return this.graphApi.request(permissionUrl)
            .catch(logErrorAndReject('Non-200 while trying to list permissions on file', this.logger))
            .then(data => {
                const permission = data.value.find(d => {
                    // unshare for public link
                    if (!email) {
                        return _.has(d, 'link.type') && !_.has(d, 'invitation');
                    }

                    // unshare for email
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
        return this.getDriveInfo().then(data => data.id);
    }

    getDriveInfo(fields = []) {
        const qs = fields.length
            ? querystring.stringify({ '$select': fields.join(',') })
            : '';

        return this.graphApi.request(`${ROOT_URL}/me/drive?${qs}`)
            .catch(logErrorAndReject('Non-200 while trying to query user details', this.logger));
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
        return super.getFileById(`${this.ROOT_URL}/me/drive/items/${fileId}`);
    }

    getFilePermissions(fileId) {
        this.logger.info('Getting OneDrive file permissions', { fileId });
        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/permissions`)
            .catch(logErrorAndReject(`Non-200 while querying file permissions ${fileId}`, this.logger))
            .then(data => data.value);
    }

    getPublicUrl(fileId) {
        return this.getFileById(fileId)
            .then(file => file.webUrl);
    }

    getPreview(fileId) {
        this.logger.info('Getting OneDrive file preview', { fileId });
        return super.getPreview(`${this.ROOT_URL}/me/drive/items/${fileId}/thumbnails`);
    }

    createFolder(folderName, parentId = rootFolderId, autorename = true) {
        parentId = parentId || rootFolderId; // there can be gotten parentId = '', which causes invalid api url
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children`;

        const body = {
            name: folderName,
            folder: { },
        };

        const autorenameField = '@microsoft.graph.conflictBehavior';
        if (autorename) {
            body[autorenameField] = "rename";
        } else {
            body[autorenameField] = "fail";
        }

        return this.graphApi.request(url, 'post', body)
            .catch(logErrorAndReject('Non-200 while trying to create folder', this.logger))
            .then(data => {
                return data.id
            });
    }

    createFileAndPopulate(fileName, content, folderId = '', useDocx = true) {
        const ext = useDocx ? 'docx' : 'doc';

        return this.graphApi.request(`https://graph.microsoft.com/v1.0/me/drive/items/${folderId}:/${fileName}.${ext}:/content`, 'put', content)
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
     * @property {string}   [resource='me/drive/root']  Folder URL, which we subscribe on
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

        const defaultPayload = {
            changeType: 'updated',
            resource: 'me/drive/root',
            expirationDateTime: addDays(new Date(), 30), // 43200 / 60 / 24 = 30 days
            clientState: '',
        };

        const fullPayload = { ...defaultPayload, ...payload };

        return this.graphApi.request(`${ROOT_URL}/subscriptions`, 'post', fullPayload)
            .catch(logErrorAndReject('Non-200 while trying to create subscription', this.logger));
    }

    /**
     * @typedef UpdateSubscriptionPayload
     * @property {Date} expirationDateTime Date, which subscription expires on (not more than 43200 hours = 30 days)
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

}

module.exports = OneDriveClient;
