const _ = require('lodash');
const GraphClient = require('./graphClient.js');
const { logErrorAndReject } = require("./util");

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
    search(
        query,
        sortProperties,
        entityTypes = ['driveItem'],
        fields = [],
        from,
        size
    ) {
        this.logger.info('Searching in Microsoft account', {
            query,
            entityTypes,
        });
        const qs = fields.length
            ? querystring.stringify({
                $select: fields.join(','),
            })
            : '';

        const requestData = {
            entityTypes,
            query: {
                queryString: query || '',
            },
            sortProperties,
            from,
            size,
        };

        for (const key in requestData) {
            if (_.isNil(requestData[key])) {
                delete requestData[key];
            }
        }

        return this.graphApi
            .request(`${this.ROOT_URL}/search/query?${qs}`, 'post', {
                requests: [requestData],
            })
            .catch(logErrorAndReject('Non-200 while searching', this.logger))
            .then(response => ({
                cursor: null, // TODO: create cursor by from and size
                items: _.get(response, 'value[0].hitsContainers[0].hits', []).map(h => h.resource),
            }));
    }
}

module.exports = SearchClient;
