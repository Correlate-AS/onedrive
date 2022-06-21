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
            ...file,
            isFolder: !!file.folder,
        }))
    };
};

const DEFAULT_SCOPES = [
    'offline_access',
    'files.readwrite'
];

function equalsToOneOfValues(value, constantsObj) {
    const keys = Object.values(constantsObj);
    return keys.includes(value);
}

function validateAndDefaultTo(value, constantsObj, defaultValue) {
    if (!equalsToOneOfValues(value, constantsObj)) {
        return defaultValue;
    }
}

/**
 * Gets param value from url
 * @param {string} url Some url
 * @param {string} paramName Param key, which value you want to get
 */
function getParamValue(url, paramName) {
    const regex = new RegExp(`^.*${paramName}=(.*?)((&.*|$))`);
    const tokenValue = url.replace(regex, '$1');

    return tokenValue;
}

module.exports = {
    logErrorAndReject,
    formatDriveResponse,
    DEFAULT_SCOPES,
    equalsToOneOfValues,
    validateAndDefaultTo,
    getParamValue,
};
