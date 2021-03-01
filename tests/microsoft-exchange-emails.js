import logger from '../lib/logger'
import User from '../lib/user';

fixture `Microsoft Exchange Email Fixture`
    .beforeEach(async t => {
        logger.init(t.testRun.test.name);
        logger.log('-----STARTED-----');
    })
    .afterEach(async t => {
        logger.log('-----FINISHED-----');
    })

    
test('Test Email Exchange', async t => {
    //Test data
    const user1 = new User('zsoltester001@zsoltester.onmicrosoft.com', 'Testuser1');
    const user2 = new User('zsoltester002@zsoltester.onmicrosoft.com', 'Testuser1');
    const message = {
        subject: 'Test Send/Receive Email' + Date.now(),
        body: '<a href=\"http://example.com\">Test URL</a>',
        recipientList: [ user2.emailAddress ]
    };

    //Step1. Send an email from one user to another with a test URL (any URL you wish) embedded
    //in the email body and specific test text in the message subject field.

    await user1.sendMail(message)
        .then(() => logger.log('Mail sent'))
        .catch(err => logger.error(err));

    let sentEmail;
    await user1.waitForEmail(message)
        .then(async email => {
            await t
                .expect(email).notEql(null)
                .expect(email.subject).eql(message.subject)
                .expect(email.body).contains(message.body)
                .expect(email.senderAddress).eql(user1.emailAddress)
                .expect(email.recipientList).eql(message.recipientList);
            sentEmail = email;
            logger.log('Sent data verified');
        })
        .catch(err => logger.error(err));

    //Step2. Verify from recipient mailbox that test URL 
    //and message subject text matches what you send in step 1.

    await user2.getMail()
        .then(async emails => {
            let receivedEmail = emails.find(e => e.uniqueId == sentEmail.uniqueId);
            await t
                .expect(receivedEmail).notEql(null)
                .expect(receivedEmail.subject).eql(sentEmail.subject)
                .expect(receivedEmail.body).contains(sentEmail.body)
                .expect(receivedEmail.senderAddress).eql(sentEmail.senderAddress)
                .expect(receivedEmail.recipientList).eql(sentEmail.recipientList);
            logger.log('Sent email content matches recipient\'s email content');
        })
        .catch(err => logger.error(err));
})


test('Test Attachment Upload', async t => {
    //Test data
    const user1 = new User('zsoltester001@zsoltester.onmicrosoft.com', 'Testuser1');
    const user2 = new User('zsoltester002@zsoltester.onmicrosoft.com', 'Testuser1');
    const message = {
        subject: 'Test Send/Receive Attachment' + Date.now(),
        body: '<h3>This is a message belonging to the attachment test</h3>',
        recipientList: [ user2.emailAddress ],
        attachments: [
            {
                name: 'The Text File',
                path: 'tests/data/text_attachment.txt'
            },
            {
                name: 'The Image file',
                path: 'tests/data/SignInActivity.JPG'
            }
        ]
    };

    //Step3. Send an email from one user to another with file attachments 
    //and specific test text in the message body field.

    await user1.sendMail(message)
        .then(() => logger.log('Mail sent'))
        .catch(err => logger.error(err));

    let sentEmail;
    await user1.waitForEmail(message)
        .then(async email => {
            await t
                .expect(email).notEql(null)
                .expect(email.subject).eql(message.subject)
                .expect(email.body).contains(message.body)
                .expect(email.senderAddress).eql(user1.emailAddress)
                .expect(email.recipientList).eql(message.recipientList);
            sentEmail = email;
            await user1.getAttachments(email.id)
                .then(async attachments => {
                    await t.expect(attachments.length).eql(message.attachments.length);
                    for (var i = 0; i < attachments.length; i++) {
                        await t.expect(attachments[i].name).eql(message.attachments[i].name);
                    }
                    sentEmail.attachments = attachments;
                });
            logger.log('Sent data verified');
        })
        .catch(err => logger.error(err));

    //Step4. Verify from recipient mailbox that attachment(s)
    //and message body text matches what you send in step 3.

    await user2.getMail()
        .then(async emails => {
            let receivedEmail = emails.find(e => e.uniqueId == sentEmail.uniqueId);
            await t.expect(receivedEmail).notEql(null)
                .expect(receivedEmail.subject).eql(sentEmail.subject)
                .expect(receivedEmail.body).contains(sentEmail.body)
                .expect(receivedEmail.senderAddress).eql(sentEmail.senderAddress)
                .expect(receivedEmail.recipientList).eql(sentEmail.recipientList);
            await user2.getAttachments(receivedEmail.id)
                .then(async attachments => {
                    await t.expect(attachments.length).eql(sentEmail.attachments.length);
                    for (var i = 0; i < attachments.length; i++) {
                        await t
                            .expect(attachments[i].name).eql(sentEmail.attachments[i].name)
                            .expect(attachments[i].size).eql(sentEmail.attachments[i].size)
                            .expect(attachments[i].contentType).eql(sentEmail.attachments[i].contentType)
                            .expect(attachments[i].contentBytes).eql(sentEmail.attachments[i].contentBytes);
                    }
                    logger.log('Sent email attachments matches received email attachments');
                });
        })
        .catch(err => logger.error(err));
})