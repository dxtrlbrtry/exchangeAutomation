import { authenticate } from './auth-service'
import { getMail, sendMail, waitForEmail, getAttachments, mailFolders } from './mail-service';

export default class {
    authToken = null;

    constructor(emailAddress, password) {
        this.emailAddress = emailAddress;
        this.password = password;
    }

    async authenticate() {
        await authenticate(this);
    }

    async getMail(mailFolder = mailFolders.INBOX, queryParams = "") {
        return await getMail(this, mailFolder, queryParams);
    }

    async sendMail(message) {
        return await sendMail(this, message);
    }

    async waitForEmail(message) {
        return await waitForEmail(this, message)
    }

    async getAttachments(id) {
        return await getAttachments(this, id);
    }
}