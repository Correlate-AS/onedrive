const querystring = require('querystring'); // deprecated for node 14-17, will be stable for node 18
const _ = require('lodash');
const GraphClient = require("./graphClient.js");
const {
    logErrorAndReject,
    formatDriveResponse,
} = require("./util");

/**
 * @typedef BaseDriveClient Base methods to work with Microsoft files
 */
class BaseDriveClient extends GraphClient {
    /**
     * @param {GraphAPI} graphApi
     * @param {object} [logger=console] Logs handler
     */
    constructor(graphApi, logger) {
        super(graphApi, logger);
    }

    /**
     * Gets drive's items from root or specified folder
     * @param {string} endpoint Service specific endpoint
     * @param {object} options
     * @returns {Promise<DriveItem[]>} https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0#properties
     */
    getFilesFrom(endpoint, options = {}) {
        const queryOptions = querystring.stringify(_.pickBy(options));
        return this.graphApi.request(`${endpoint}?${queryOptions}`)
            .catch(logErrorAndReject(`Non-200 while querying folder: ${endpoint}`, this.logger))
            .then(formatDriveResponse);
    }

    /**
     * Gets drive item by its id
     * @param {string} endpoint Service specific endpoint
     * @param {object} options
     * @returns {Promise<DriveItem>} https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0#properties
     * @async
     */
    getFileById(endpoint, options) {
        this.logger.info(`Query options`, { options });
        const queryOptions = this._generateQueryOptions(options);

        return this.graphApi
            .request(`${endpoint}?${queryOptions}`)
            .catch(
                logErrorAndReject(`Non-200 while querying file ${endpoint}`, this.logger)
            )
            .then((data) => {
                return {
                    ...data,
                    packageType: data.package ? data.package.type : "",
                };
            });
    }

    /**
     * Gets Preview data
     * @param {string} endpoint Service specific endpoint
     * @returns Preview data
     * @async
     */
    getPreview(endpoint) {
        return this.graphApi.request(endpoint)
            .catch(logErrorAndReject(`Non-200 while querying file ${endpoint}`, this.logger))
            .then(data => {
                return data.value;
            });
    }

    /**
     * Creates sharing link of specified type if it does not already exist
     * or returns sharing link if link of such a type already exists
     * @param {string} endpoint Service specific endpoint https://learn.microsoft.com/en-us/graph/api/driveitem-createlink?view=graph-rest-1.0&tabs=http#http-request
     * @returns {Promise<Permission>} https://learn.microsoft.com/en-us/graph/api/resources/permission?view=graph-rest-1.0#properties
     * @async
     */
    createSharingLink(endpoint) {
        const body = {
            type: 'view',
            scope: 'users'
        };

        return this.graphApi.request(endpoint, 'post', body)
            .catch(logErrorAndReject(`Non-200 while creating sharing link ${endpoint}`, this.logger));
    }

    /**
     * Grants access to file, which is represented by link, for specified emails
     * @param {string} link Link, which is related to file to be shared
     * @param {string[]} emails Emails, which are going to get access to file
     * @returns {Promise<Permission>} Graph API Permission but `grantedToIdentities` and `grantedToIdentitiesV2` can be not updated
     * https://learn.microsoft.com/en-us/graph/api/resources/permission?view=graph-rest-1.0#properties
     * @async
     */
    grantAccessForLink(link, emails = []) {
        this.logger.info(`Granting access for link`, { link, emails });

        const encodedSharingUrl = this._encodeSharingUrl(link);
        const endpoint = `${this.ROOT_URL}/shares/${encodedSharingUrl}/permission/grant`;
        const body = {
            recipients: emails.map(e => ({ email: e })),
            roles: ['read'],
        };

        return this.graphApi.request(endpoint, 'post', body)
            .catch(logErrorAndReject(`Non-200 while granting access for link ${link}`, this.logger))
            .then(data => {
                return data.value;
            });
    }

    /**
     * Deletes file's permission
     * @param {string} endpoint Service specific endpoint https://learn.microsoft.com/en-us/graph/api/permission-delete?view=graph-rest-1.0&tabs=http#http-request
     * @returns {Promise<undefined>}
     * @async
     */
    deletePermissions(endpoint) {
        return this.graphApi.request(endpoint, 'delete')
            .catch(logErrorAndReject(`Non-200 while listing permissions ${endpoint}`, this.logger))
            .then(() => {});
    }

    /**
     * Generates valid Onedrive and Sharepoint query options
     * https://learn.microsoft.com/en-us/graph/query-parameters
     * @param {object} options
     * @param {string[]} fields
     * @param {string[]} expand Expand relations
     * @returns {string}
     */
    _generateQueryOptions(options = {}) {
        const { fields = [], expand = [] } = options;
        let optionValid = [];

        /**
         * Expand can affect `$select`, so it have to go first
         */
        if (expand.length) {
          optionValid.push(querystring.stringify({ $expand: expand.join(",") }));
        }

        if (fields.length) {
          // https://stackoverflow.com/a/44571731/13745132
          optionValid.push(querystring.stringify({ $select: [...fields, "id"].join(",") }) );
        }

        return optionValid.join("&");
    }

    /**
     * Generates encoded url, which is necessary for sharing API
     * https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0&tabs=http#encoding-sharing-urls
     * @param {string} url
     * @returns {string}
     */
    _encodeSharingUrl(url) {
        const encodedToBase64 = Buffer.from(url).toString('base64');
        const encodedWithValidChars = encodedToBase64.replace(/=+$/,'').replace('/', '_').replace('+', '-');
        return 'u!' + encodedWithValidChars;
    }

    /**
     * Returns provided containerId or if it falsy value - id of root container
     * @param {string} containerId Container ID, where container can be drive, site, folder etc.
     * @returns {string}
     */
    _validateContainer(containerId) {
        return containerId || this.ROOT_FOLDER;
    }
}

module.exports = BaseDriveClient;