const _ = require('lodash');
const GraphClient = require('./graphClient.js');
const { logErrorAndReject, removeNilValues } = require("./util");

/**
 * For some reason there are duplicates in search result if size is over 25.
 * Option trimDuplicates = false makes no changes.
 * 
 * These duplicates are counted in 'total'.
 * If, e.g., we have 3 uniq hits and 2 duplicates then for page size = 3, total = 5, BUT moreResultsAvailable = false.
 * Others are curious too: https://learn.microsoft.com/en-us/answers/questions/899486/microsoft-search-api-response-incorrect-39total39.html
 */
const MAX_RESULTS = 25;

/**
 * Microsoft Search API
 */

class SearchClient extends GraphClient {
    /**
     * 
     * @param {import("./graphApi.js").GraphAPI} graphApi 
     * @param {object} [logger] Logs handler
     */
    constructor(graphApi, logger) {
        super(graphApi, logger);
    }

    /**
     * @typedef {object} Response 
     * @property {string} [cursor] Next page pointer
     * @property {object[]} [items=[]] Found items
     */

    /**
     * Searches across microsoft entities
     * @param {string} query Specify term [and filters], filters (https://learn.microsoft.com/en-us/graph/search-concept-files#example-5-use-filters-in-search-queries)
     * @param {string} [sortProperties] Coma separated sort options (e.g., `name asc, createdDateTime desc`)
     * @param {sting[]} [entityTypes=['driveItem']] Type of items to be searched. Possible values are: `list`, `site`, `listItem`, `message`, `event`, `drive`, `driveItem`, `externalItem`
     * @param {sting[]} [fields=[]] Fields of entity you search
     * @param {sting} [cursor] Next page pointer
     * @param {number} [maxResults=MAX_RESULTS] Max results per page
     * @returns {Promise<Response>}
     */
    search({
        query,
        sortProperties,
        entityTypes = ['driveItem'],
        fields,
        cursor,
        maxResults = MAX_RESULTS,
    }) {
        let requestData = {
            entityTypes,
            query: {
                queryString: query,
            },
            sortProperties: _parseSortProperties(sortProperties),
            from: 0,
            size: maxResults > MAX_RESULTS ? MAX_RESULTS : maxResults,
            fields,
        };
        requestData = _adjustPage(requestData, cursor);
        requestData = removeNilValues(requestData);

        this.logger.info('Searching in Microsoft account', { ...requestData });

        return this.graphApi
            .request(`${this.ROOT_URL}/search/query`, 'post', { requests: [requestData] })
            .catch(logErrorAndReject('Non-200 while searching', this.logger))
            .then((response) => _formatResponse(requestData, response));
    }
}

/**
 * @typedef SortProperty Microsoft Search API acceptable sort property
 * @type {object}
 * @property {string} name
 * @property {boolean} isDescending
 */

/**
 * Parses sort option into acceptable ones for Microsoft Search API
 * @param {string} [sortProperties] Coma separated sort options (e.g., `name asc, createdDateTime desc`)
 * @returns {SortProperty[]}
 */
function _parseSortProperties(sortProperties = '') {
    const sortings = sortProperties.split(', ');

    return sortings.map(sorting => {
        const values = sorting.split(' ');
        
        return {
            name: values[0],
            isDescending: (values[1] || '').toUpperCase() === 'DESC',
        };
    });
}

/**
 * @typedef MicrosoftSearchRequest Microsoft Search API request
 * @type {object}
 * @property {sting[]} entityTypes Type of items to be searched. Possible values are: `list`, `site`, `listItem`, `message`, `event`, `drive`, `driveItem`, `externalItem`
 * @property {object} query Contains the query terms
 * @property {object} query.queryString The search query containing the search terms
 * @property {SortProperty} [sortProperties] Indicates the order to sort search results
 * @property {number} [from] The size of the page to be retrieved.The maximum value is 1000
 * @property {number} [size] Specifies the offset for the search results. Offset 0 returns the very first result
 */

/**
 * Returns request data according to the page cursor pointed to
 * @param {MicrosoftSearchRequest} requestData 
 * @param {string} cursor Page pointer
 * @returns {MicrosoftSearchRequest}
 */
function _adjustPage(requestData, cursor) {
    const decodedCursor = _decodeCursor(cursor);

    if (!decodedCursor) {
        return requestData;
    }
    
    return {
        ...requestData,
        ...decodedCursor,
    };
}

/**
 * Encodes cursor and returns data hidden in it
 * @param {string} cursor Base64 string
 * @returns {object} Encoded cursor
 */
function _encodeCursor(cursor) {
    return cursor ? Buffer.from(JSON.stringify(cursor)).toString('base64') : null;
}

/**
 * Decodes next page info into string
 * @param {object} data Next page data
 * @returns {string} Decoded cursor
 */
function _decodeCursor(data) {
    return data && JSON.parse(Buffer.from((data), 'base64').toString('ascii'));
}

/**
 * @typedef {object} MicrosoftSearchResponse Microsoft Search API response (https://learn.microsoft.com/en-us/graph/api/resources/searchresponse?view=graph-rest-1.0)
 */

/**
 * Formats response into the view other client returns
 * @param {MicrosoftSearchRequest} requestData 
 * @param {MicrosoftSearchResponse} response 
 * @returns {Response}
 */
function _formatResponse(requestData, response) {
    const result =  _.get(response, 'value[0].hitsContainers[0]');
    const isNextPage = _.get(result, 'moreResultsAvailable');
    const hits = _.get(result, 'hits', []);

    return {
        cursor: isNextPage
            ? _encodeCursor({ // next page
                ...requestData,
                from: requestData.from + requestData.size,
            })
            : null,
        items: hits.map(h => h.resource)
    };
}

module.exports = SearchClient;
