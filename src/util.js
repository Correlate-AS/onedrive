const _ = require('lodash');
const querystring = require('querystring');

const logErrorAndReject = (text, logger) => {
    return error => {
        if ('response' in error) {
            const { response } = error;
            logger.error(_.pick(response, ['status', 'data', 'config']), text);
            return Promise.reject(error);
        }
        logger.error('Unexpected error', { err: error });
        return Promise.reject(error);
    }
};

const formatDriveResponse = data => {
    const { value } = data;
    const nextLink = data['@odata.nextLink'];
    let skiptoken = null;
    if (nextLink) {
        const nextLinkQuery = new URL(nextLink).search;
        const params = querystring.parse(nextLinkQuery);
        skiptoken = params['$skiptoken'] || null;
    }
    return {
        cursor: skiptoken,
        items: value.map(file => ({
            id: file.id,
            isFolder: !!file.folder,
            name: file.name
        }))
    };
};

const DEFAULT_SCOPES = [
    'offline_access',
    'files.readwrite'
];

module.exports = {
    logErrorAndReject,
    formatDriveResponse,
    DEFAULT_SCOPES
};
