import * as authService from './auth-service'
import * as mailService from './mail-service';

export default class {

    constructor(emailAddress, password) {
        this.emailAddress = emailAddress;
        this.password = password;
        this.authToken = null;
    }

    async authenticate(expectedStatusCode) {
        this.authToken = await authService.authenticate(this, expectedStatusCode);
    }

    async sendMail(message, expectedStatusCode) {
        await mailService.sendMail(this, message, expectedStatusCode);
    }

    async getMail(mailFolder, queryParams, expectedStatusCode) {
        return await mailService.getMail(this, mailFolder, queryParams, expectedStatusCode);
    }

    async waitForEmail(message, expectedStatusCode) {
        return await mailService.waitForEmail(this, message, expectedStatusCode)
    }

    async getAttachments(id, expectedStatusCode) {
        return await mailService.getAttachments(this, id, expectedStatusCode);
    }
}