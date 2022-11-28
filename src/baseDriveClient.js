const querystring = require('querystring'); // deprecated for node 14-17, will be stable for node 18
const GraphClient = require("./graphClient.js");
const {
    logErrorAndReject,
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
     * Gets drive item by its id
     * @param {*} endpoint
     * @param {*} options
     * @returns {Promise<driveItem>} https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0#properties
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
     * Shares drive item for email
     * @param {string} endpoint Service specific endpoint
     * @param {string} email
     * @returns {Promise}
     * @async
     */
    shareForEmail(endpoint, email) {
        return this.graphApi.request(endpoint, 'POST', {
                requireSignin: true,
                sendInvitation: false,
                roles: ["read"],
                recipients: [{
                    email,
                }],
                message: "File shared through Correlate"
            })
            .catch(logErrorAndReject('Non-200 while trying to share file', this.logger));
    }

    /**
     * Unshares drive item for email
     * @param {string} endpoint Service specific endpoint
     * @returns {Promise}
     * @async
     */
    unshareFrom(endpoint) {
        return this.graphApi.request(endpoint, "DELETE")
            .catch(logErrorAndReject('Non-200 while removing permission', this.logger));
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
}

module.exports = BaseDriveClient;