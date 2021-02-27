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
        this.messageId = email['internetMessageId'];
        this.senderAddress = email['sender']['emailAddress']['address'];
        this.attachments = [];
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

export async function getMail(user, mailFolder, queryParams) {
    if (!user.isAuthenticated()) {
        await user.authenticate();
    }

    const options = { 
        method: 'GET',
        url: 'https://graph.microsoft.com/v1.0/me/mailFolders/' + mailFolder + '/messages' + queryParams,
        headers: { 'Authorization': user.authToken },
        json: false,
        resolveWithFullResponse: true
    };

    const response = await request(options)
        .catch(err => {
            console.log(err);
        });
    
    const emailList = JSON.parse(response.body)['value'];
    for (var i = 0; i < emailList.length; i++) {
        user.emails.push(new Email(emailList[i]))
    }

    return response;
}

export async function sendMail(user, message) {
    if (!user.isAuthenticated()) {
        await user.authenticate();
    }

    const email = await composeEmail(message);
    const options = { 
        method: 'POST',
        url: 'https://graph.microsoft.com/v1.0/me/sendMail',
        headers: { 'Authorization': user.authToken, 'Content-Type': 'application/json' },
        body: email,
        json: false,
        resolveWithFullResponse: true
    };

    return await request(options)
        .catch(err => {
            console.log("Error Sending Email: " + err);
        });
}

export async function waitForEmail(user, message) {
    const maxAttempts = 3;
    const attemptInterval = 2000;
    
    let attempts = 0;
    let lastSentEmail = null;
    let response;
    
    while(lastSentEmail == null || !lastSentEmail.body.includes(message.body)) {
        if (attempts < maxAttempts){
            attempts++;
        } else {
            throw 'Timed out while validating sent mail - ' + lastSentEmail.body + ' does not include substring ' + message.body;
        }
        await new Promise(r => setTimeout(r, attemptInterval));
        response = await user.getMail(mailFolders.SENTITEMS, '?$top=1');
        if (response.statusCode == 200) {
            lastSentEmail = user.emails[0];
        }
    }

    return response;
}

export async function getAttachments(user, id) {
    if (!user.isAuthenticated()) {
        await user.authenticate();
    }

    const options = { 
        method: 'GET',
        url: 'https://graph.microsoft.com/v1.0/me/messages/' + id + '/attachments',
        headers: { 'Authorization': user.authToken },
        json: false,
        resolveWithFullResponse: true
    };

    const response = await request(options)
        .catch(err => {
            console.log("Error Sending Email: " + err);
        });

    const attachments = JSON.parse(response.body)['value'];
    for (var i = 0; i < user.emails.length; i++) {
        for (var j = 0; j < attachments.length; j++) {
            if (user.emails[i].id == id) {
                user.emails[i].attachments.push(new Attachment(attachments[j]));
            }
        }
    }

    return response;
}

async function composeEmail(message) {
    const email = {
        "message": {
            "subject": message.subject,
            "body": {
                "contentType":"HTML",
                "content": message.body
            },
            "toRecipients": [],
            'attachments': []
        }
    }
    for (var i = 0; i < message.recipientList.length; i++) {
        email['message']['toRecipients'].push({ 'emailAddress': { 'address': message.recipientList[i] }});
    }
    for (var i = 0; i < message.attachments.length; i++) {
        let attachment = message.attachments[i];
        let file = new File(attachment.path);
        let byteArray = await fileToByteConverter(file);
        let contentBytes = byteArray.match('(?<=\,).*').map(val => { return val })[0];
        let contentType = file.type;

        email['message']['attachments'].push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": attachment.name,
            "contentType": contentType,
            "contentBytes": contentBytes});
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