/**
 * @typedef GraphClient Base methods to work with Microsoft Graph API
 */
 class GraphClient {
    /**
     * @param {GraphAPI} graphApi
     * @param {object} [logger=console] Logs handler
     */
  constructor(graphApi, logger) {
    this.graphApi = graphApi;
    this.logger = logger || console;
    this.ROOT_URL = "https://graph.microsoft.com/v1.0";
  }
}

module.exports = GraphClient;