import logger from './logger';
import { validator, schemas } from './schema-validator';
import { baseUrl } from './constants';
const request = require("request-promise");
const FileReader = require("filereader");
const File = require('file');

export const mailFolders = {
    INBOX: 'inbox',
    SENTITEMS: 'sentItems'
};

class Email {
    constructor(email) {
        this.id = email['id'];
        this.subject = email['subject'];
        this.body = email['body']['content'].match('(?<=<body>)(.*)(?=<\/body>)').map(val => { return val })[0];;
        this.uniqueId = email['internetMessageId'];
        this.senderAddress = email['sender']['emailAddress']['address'];
        this.recipientList = [];
        for (var i = 0; i < email['toRecipients'].length; i++) {
            this.recipientList.push(email['toRecipients'][i]['emailAddress']['address']);
        }
    }
}

class Attachment {
    constructor(attachment) {
        this.name = attachment['name'];
        this.size = attachment['size'];
        this.contentType = attachment['contentType'];
        this.contentBytes = attachment['contentBytes'];
    }
}

export async function getMail(user, mailFolder, queryParams, expectedStatusCode) {
    const options = { 
        method: 'GET',
        url: baseUrl + '/v1.0/me/mailFolders/' + mailFolder + '/messages' + queryParams,
        headers: { 'Authorization': user.authToken },
        json: false,
        resolveWithFullResponse: true
    };

    return await request(options, expectedStatusCode)
        .then(resp => {
            if (resp.statusCode != expectedStatusCode) {
                throw 'Expected status code is ' + expectedStatusCode + ' but got ' + resp.statusCode;
            } else if (resp.statusCode == 200) {
                validator.validate(resp.body, schemas.definitions.getMail);

                let emails = [];
                let emailList = JSON.parse(resp.body)['value'];
                for (var i = 0; i < emailList.length; i++) {
                    emails.push(new Email(emailList[i]))
                }

                return emails;
            }
        
        })
        .catch(err => logger.error('Error getting mailbox: ' + err));
}

export async function sendMail(user, message, expectedStatusCode) {
    const email = await composeEmail(message);
    const options = { 
        method: 'POST',
        url: baseUrl + '/v1.0/me/sendMail',
        headers: { 'Authorization': user.authToken, 'Content-Type': 'application/json' },
        body: email,
        json: false,
        resolveWithFullResponse: true
    };

    await request(options, expectedStatusCode)
        .then(resp => {
            if (resp.statusCode != expectedStatusCode) {
                throw 'Expected status code is ' + expectedStatusCode + ' but got ' + resp.statusCode;
            } else if (resp.statusCode == 202) { }
        })
        .catch(err => logger.error("Error Sending Email: " + err));
}

export async function waitForEmail(user, message) {
    const maxAttempts = 5;
    const attemptInterval = 2000;
    
    let attempts = 0;
    let lastSentEmail = null;
    
    try {
        while(lastSentEmail == null || !lastSentEmail.body.includes(message.body)) {
            if (attempts <= maxAttempts) {
                attempts++;
            } else {
                throw 'Timed out while validating sent mail body - ' + lastSentEmail.body + ' does not include substring ' + message.body;
            }
            await new Promise(r => setTimeout(r, attemptInterval));
            await user.getMail(mailFolders.SENTITEMS, '?$top=5&orderby=receivedDateTime%20desc', 200)
                .then(emails => {
                    if (emails.length > 0) {
                        lastSentEmail = emails[0];
                    }
                });
        }
    }
    catch(err) { logger.error('Error while waiting for mail: ' + err) }

    return lastSentEmail;
}

export async function getAttachments(user, id, expectedStatusCode) {
    const options = { 
        method: 'GET',
        url: baseUrl + '/v1.0/me/messages/' + id + '/attachments',
        headers: { 'Authorization': user.authToken },
        json: false,
        resolveWithFullResponse: true
    };

    return await request(options, expectedStatusCode)
        .then(resp => {
            if (resp.statusCode != expectedStatusCode) {
                throw 'Expected status code is ' + expectedStatusCode + ' but got ' + resp.statusCode;
            } else if (resp.statusCode == 200) {
                validator.validate(resp.body, schemas.definitions.getAttachments);
                let attachments = [];
                const attachmentsList = JSON.parse(resp.body)['value'];
                for (var i = 0; i < attachmentsList.length; i++) {
                    attachments.push(new Attachment(attachmentsList[i]));
                }
                return attachments;;
            }
        })
        .catch(err => logger.error("Error getting attachments for email Id " + id + ": " + err));
}

async function composeEmail(message) {
    const email = {
        "message": {
            "subject": message.subject,
            "body": {
                "contentType":"HTML",
                "content": message.body
            },
            "toRecipients": []
        }
    }
    for (var i = 0; i < message.recipientList.length; i++) {
        email['message']['toRecipients'].push({ 'emailAddress': { 'address': message.recipientList[i] }});
    }
    if (message.attachments != undefined && message.attachments.length > 0)
    {
        email['message']['attachments'] = [];
        for (var i = 0; i < message.attachments.length; i++) {
            try {        
                let attachment = message.attachments[i];
                let file = new File(attachment.path);
                let byteArray = await fileToByteConverter(file);
                let contentBytes = byteArray.match('(?<=\,).*').map(val => { return val })[0];
    
                email['message']['attachments']
                    .push({
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": attachment.name,
                        "contentType": file.type,
                        "contentBytes": contentBytes
                    });
            } catch(err) { logger.error(err); }
        }
    }
    
    return JSON.stringify(email);
}

function fileToByteConverter(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}