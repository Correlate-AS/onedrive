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

    getFileById(endpoint) {
        return this.graphApi
            .request(endpoint)
            .catch(
                logErrorAndReject(`Non-200 while querying file ${endpoint}`, this.logger)
            )
            .then((data) => {
                return {
                    name: data.name,
                    webUrl: data.webUrl,
                    packageType: data.package ? data.package.type : "",
                    parentReference: data.parentReference,
                };
            });
    }

    getPreview(endpoint) {
        return this.graphApi.request(endpoint)
            .catch(logErrorAndReject(`Non-200 while querying file ${endpoint}`, this.logger))
            .then(data => {
                return  data.value;
            });
    }
}

module.exports = BaseDriveClient;