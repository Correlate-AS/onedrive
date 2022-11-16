const _ = require('lodash');
const GraphClient = require('./graphClient.js');
const { logErrorAndReject } = require("./util");

/**
 * For some reason there are duplicates in search result if size is over 25.
 * Option trimDuplicates = false makes no changes.
 */
const MAX_RESULTS = 25;

/**
 * Microsoft Search API
 */

class SearchClient extends GraphClient {
    constructor(graphApi, logger) {
        super(graphApi, logger);
    }

    /**
     *
     * @param {*} query Specify term, filters (https://learn.microsoft.com/en-us/graph/search-concept-files#example-5-use-filters-in-search-queries)
     * @param {*} [sortProperties]
     * @param {string} [sortProperties.name] You can be certain to use `createdDateTime`, `name`, `lastModifiedDateTime`, others should be checked
     * @param {boolean} [sortProperties.isDescending] Sort order (ASC -> `false`, DESC -> `true`)
     * @param {sting[]} [entityTypes=['driveItem']] Possible values are: `list`, `site`, `listItem`, `message`, `event`, `drive`, `driveItem`, `externalItem`
     * @param {sting[]} [fields=[]] Fields of entity you search
     * @returns
     */
    search({
        query = '',
        sortProperties,
        entityTypes = ['driveItem'],
        fields = [],
        cursor,
        maxResults = MAX_RESULTS,
    }) {
        let requestData = {
            entityTypes,
            query: {
                queryString: query || '',
            },
            sortProperties: parseSortProperties(sortProperties),
            from: 0,
            size: maxResults > MAX_RESULTS ? MAX_RESULTS : maxResults,
        };

        const decodedCursor = _decodeCursor(cursor);
        if (decodedCursor) {
            requestData = {
                entityTypes,
                query: requestData.query,
                sortProperties: requestData.sortProperties,
                ...decodedCursor,
            }
        }

        this.logger.info('Searching in Microsoft account', { ...requestData });

        requestData = _removeNilValues(requestData);

        const qs = generateQueryParams({ fields });

        return this.graphApi
            .request(`${this.ROOT_URL}/search/query?${qs}`, 'post', {
                requests: [requestData],
            })
            .catch(logErrorAndReject('Non-200 while searching', this.logger))
            .then((response) => formatResponse(requestData, response));
    }
}

function parseSortProperties(sortProperties) {
    const sortings = sortProperties.split(', ');

    return sortings.map(sorting => {
        const values = sorting.split(' ');
        
        return {
            name: values[0],
            isDescending: values[1].toUpperCase() === 'DESC',
        };
    });
}

function generateQueryParams({ fields = [] }) {
    const qs = fields.length
        ? querystring.stringify({
            $select: fields.join(','),
        })
        : '';

    return qs;
}

function _removeNilValues(obj) {
    const newObj = { ...obj };

    for (const key in newObj) {
        if (_.isNil(newObj[key])) {
            delete newObj[key];
        }
    }

    return newObj;
}

function _decodeCursor(cursor) {
    return cursor ? JSON.parse(cursor) : null;
}

function _encodeCursor(data) {
    return JSON.stringify(data)
}

function formatResponse(requestData, response) {
    return {
        cursor: _.get(response, 'value[0].hitsContainers[0].moreResultsAvailable')
            ? _encodeCursor({ // next page
                size: requestData.size,
                from: requestData.from + requestData.size,
            })
            : null,
        items: _.get(response, 'value[0].hitsContainers[0].hits', []).map(h => h.resource)
    };
}

module.exports = SearchClient;
