// More details here: https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_put_content?view=odsp-graph-online#conflict-resolution-behavior
const UPLOAD_CONFLICT_RESOLUTION_MODES = {
    FAIL: 'fail',
    REPLACE: 'replace',
    RENAME: 'rename',
}

module.exports = {
    UPLOAD_CONFLICT_RESOLUTION_MODES
}
