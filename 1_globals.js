// set these according to your Zoom App
const ACCOUNT_ID = 'q8NUOjMPT0-V6NlQXzPE4A';
const CLIENT_ID = '_F_cBZz7SeugyGYAsBtvg';
const CLIENT_SECRET = 'Q05iEQ0adlyGkU4CNySZJio2CAYvRaF6';

function getZoomAccessToken() {
    const tokenUrl = 'https://zoom.us/oauth/token';
    const basic = Utilities.base64Encode(`${CLIENT_ID}:${CLIENT_SECRET}`);
    const options = {
        method: 'post',
        headers: {
            Authorization: 'Basic ' + basic,
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: {
            grant_type: 'account_credentials',
            account_id: ACCOUNT_ID
        },
        muteHttpExceptions: true
    };
    const resp = UrlFetchApp.fetch(tokenUrl, options);
    if (resp.getResponseCode() !== 200) {
        throw new Error('Token request failed: ' + resp.getContentText());
    }
    const data = JSON.parse(resp.getContentText());
    // the below log()s are purely for making sure there is a token, keep it commented out otherwise
    // Logger.log(resp.getResponseCode());
    // Logger.log(resp.getContentText());
    return data.access_token;
}
