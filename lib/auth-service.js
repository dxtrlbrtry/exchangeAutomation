const request = require("request-promise");
import logger from './logger';
import { validator, schemas } from './schema-validator';

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
        },
        resolveWithFullResponse: true
    };

    await request(options)
        .then(resp => {
            if (resp.statusCode != 200) { throw 'Status code was not OK'; }
            validator.validate(resp.body, schemas.definitions.authentication)
            user.authToken = JSON.parse(resp.body)['access_token'];
            logger.log('Signed in as ' + user.emailAddress);
        })
        .catch(err => logger.error(err));
}
