const _ = require('lodash');

const logErrorAndReject = (text, logger) => {
    return error => {
        if ('response' in error) {
            const { response } = error;
            logger.error(text, _.pick(response, ['status', 'data', 'config.url']));
            return Promise.reject(new Error(text));
        }
        logger.error('Unexpected error', { err: error });
        return Promise.reject(error);
    }
};

const DEFAULT_SCOPES = [
    'offline_access',
    'openid',
    'user.read',
    'files.readwrite'
];

module.exports = {
    logErrorAndReject,
    DEFAULT_SCOPES
};
