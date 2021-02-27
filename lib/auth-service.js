const request = require("request-promise");

export async function authenticate(user) {
    const options = { 
        method: 'POST',
        url: 'https://login.microsoftonline.com/7ce4c09e-ec00-4d0a-ba6c-f11abd0c074b/oauth2/v2.0/token',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        form: { 
            client_id: 'ae6f7b09-5929-42dd-8114-4f99e851e09b',
            scope: 'https://graph.microsoft.com/.default',
            client_secret: 'op8IHbpWdxeO1ES.Bg20z._D04p_zDDfh8',
            username: user.emailAddress,
            password: user.password,
            grant_type: 'password'
        }
    };

    await request(options)
        .then(resp => {
            user.authToken = JSON.parse(resp)['access_token'];
        })
}