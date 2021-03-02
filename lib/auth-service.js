const request = require("request-promise");
import logger from './logger';
import { validator, schemas } from './schema-validator';
import { loginUrl, baseUrl, client_id, client_secret, tenant_id } from './constants';

export async function authenticate(user, expectedStatusCode) {
    const options = { 
        method: 'POST',
        url: loginUrl + '/' + tenant_id + '/oauth2/v2.0/token',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        form: { 
            client_id: client_id,
            scope: baseUrl + '/.default',
            client_secret: client_secret,
            username: user.emailAddress,
            password: user.password,
            grant_type: 'password'
        },
        resolveWithFullResponse: true
    };

    return await request(options, expectedStatusCode)
        .then(resp => {
            if (resp.statusCode != expectedStatusCode) {
                throw 'Expected status code is ' + expectedStatusCode + ' but got ' + resp.statusCode;
            } else if (resp.statusCode == 200) {
                validator.validate(resp.body, schemas.definitions.authentication);
                logger.log('Signed in as ' + user.emailAddress);
                return JSON.parse(resp.body)['access_token'];
            }
        })
        .catch(err => logger.error(err));
}
