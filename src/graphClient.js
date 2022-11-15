class GraphClient {
  constructor(graphApi, logger) {
    this.graphApi = graphApi;
    this.logger = logger;
    this.ROOT_URL = "https://graph.microsoft.com/v1.0";
  }
}

module.exports = GraphClient;