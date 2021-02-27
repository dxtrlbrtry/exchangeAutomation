import { Logger } from '../lib/logger'
import { User } from '../lib/user';

let logger;

fixture `Microsoft Exchange Email Fixture`
    .beforeEach(async t => {
        logger = new Logger(t.testRun.test.name);
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
        recipientList: [ user2.emailAddress ],
        attachments: []
    };

    //Step1. Send an email from one user to another with a test URL (any URL you wish) embedded
    //in the email body and specific test text in the message subject field.

    await user1.sendMail(message)
        .then(async resp =>
            await t.expect(resp.statusCode).eql(202));
    logger.log('Mail sent');

    let sentEmail;
    await user1.waitForEmail(message)
        .then(async resp => {
            sentEmail = user1.emails[0];
            await t
                .expect(resp.statusCode).eql(200)
                .expect(sentEmail).notEql(null)
                .expect(sentEmail.subject).eql(message.subject)
                .expect(sentEmail.body).contains(message.body)
                .expect(sentEmail.senderAddress).eql(user1.emailAddress);
        });
  
    logger.log('Sent data verified');

    //Step2. Verify from recipient mailbox that test URL 
    //and message subject text matches what you send in step 1.

    await user2.getMail()
        .then(async resp => 
            await t.expect(resp.statusCode).eql(200));
    logger.log('Retrieved User2 mailbox');

    let receivedEmail = user2.emails.find(e => e.emailId == sentEmail.emailId);
    await t
        .expect(receivedEmail).notEql(null)
        .expect(receivedEmail.subject).eql(message.subject)
        .expect(receivedEmail.body).contains(message.body)
        .expect(receivedEmail.senderAddress).eql(user1.emailAddress);
    logger.log('Sent email content matches recipient\'s email content');
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
        .then(async resp => 
            await t.expect(resp.statusCode).eql(202));
    logger.log('Mail sent');

    let sentEmail;
    await user1.waitForEmail(message)
        .then(async resp => {
            sentEmail = user1.emails[0];
            await t
                .expect(resp.statusCode).eql(200)
                .expect(sentEmail).notEql(null)
                .expect(sentEmail.subject).eql(message.subject)
                .expect(sentEmail.body).contains(message.body)
                .expect(sentEmail.senderAddress).eql(user1.emailAddress);
        });
    
    await user1.getAttachments(sentEmail.id)
        .then(async resp => {
            await t
                .expect(resp.statusCode).eql(200)
                .expect(sentEmail.attachments.length).eql(message.attachments.length);
            for (var i = 0; i < sentEmail.attachments.length; i++) {
                await t.expect(sentEmail.attachments[i].name).eql(message.attachments[i].name);
            }
        });  
    logger.log('Sent data verified');

    //Step4. Verify from recipient mailbox that attachment(s)
    //and message body text matches what you send in step 3.

    await user2.getMail()
        .then(async resp => 
            await t.expect(resp.statusCode).eql(200));
    logger.log('Get User2 mailbox');

    let receivedEmail = user2.emails.find(e => e.emailId == sentEmail.emailId);
    await t.expect(receivedEmail).notEql(null);

    await user2.getAttachments(receivedEmail.id)
        .then(async resp => {
            await t
                .expect(resp.statusCode).eql(200)
                .expect(receivedEmail.attachments.length).eql(sentEmail.attachments.length);
            for (var i = 0; i < receivedEmail.attachments.length; i++) {
                await t
                    .expect(receivedEmail.attachments[i].name).eql(sentEmail.attachments[i].name)
                    .expect(receivedEmail.attachments[i].size).eql(sentEmail.attachments[i].size)
                    .expect(receivedEmail.attachments[i].contentType).eql(sentEmail.attachments[i].contentType)
                    .expect(receivedEmail.attachments[i].contentBytes).eql(sentEmail.attachments[i].contentBytes);
            }
        });
    logger.log('Sent email attachments matches received email attachments');
})