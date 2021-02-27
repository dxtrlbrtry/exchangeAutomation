import { Logger } from '../lib/logger'
import { User } from '../lib/user';

fixture `Microsoft Exchange Email Fixture`

test('Test Email Exchange', async t => {
    const logger = new Logger('Test Email Exchange');
    logger.log('-----STARTED-----');
    //Test data
    const user1 = new User('zsoltester001@zsoltester.onmicrosoft.com', 'Testuser1');
    const user2 = new User('zsoltester002@zsoltester.onmicrosoft.com', 'Testuser1');
    const message = {
        subject: 'Test Send/Receive Email' + Date.now(),
        body: '<a href=\"http://example.com\">Test URL</a>',
        recipientList: [ user2.emailAddress ],
        attachments: []
    };

    //Send the email
    var response = await user1.sendMail(message);
    await t.expect(response.statusCode).eql(202);
    logger.log('Mail sent');

    //Wait for the email to arrive
    var response = await user1.waitForEmail(message);
    let sentEmail = user1.emails[0];
    await t
        .expect(response.statusCode).eql(200)
        .expect(sentEmail).notEql(null)
        .expect(sentEmail.subject).eql(message.subject)
        .expect(sentEmail.body).contains(message.body)
        .expect(sentEmail.senderAddress).eql(user1.emailAddress);
    logger.log('Verify sent data');

    //Verify that the email has arrived
    var response = await user2.getMail();
    await t.expect(response.statusCode).eql(200);
    logger.log('Retrieved mailbox');

    //Verify that the received email data matches the sent email data
    let receivedEmail = user2.emails.find(em => em.emailId == sentEmail.emailId);
    await t
        .expect(receivedEmail).notEql(null)
        .expect(receivedEmail.subject).eql(message.subject)
        .expect(receivedEmail.body).contains(message.body)
        .expect(receivedEmail.senderAddress).eql(user1.emailAddress);
    logger.log('Sent email matches recipient\'s email');
    logger.log('-----PASSED-----');
})


test('Test Attachment Upload', async t => {
    const logger = new Logger('Test Attachment Upload');
    logger.log('-----STARTED-----');
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

    //Send the email
    var response = await user1.sendMail(message);
    await t.expect(response.statusCode).eql(202);
    logger.log('Mail sent');

    //Wait for the email to arrive
    var response = await user1.waitForEmail(message);
    let sentEmail = user1.emails[0];
    await t
        .expect(response.statusCode).eql(200)
        .expect(sentEmail).notEql(null)
        .expect(sentEmail.subject).eql(message.subject)
        .expect(sentEmail.body).contains(message.body)
        .expect(sentEmail.senderAddress).eql(user1.emailAddress);

    //Save attachment data in user
    var response = await user1.getAttachments(sentEmail.id);
    await t.expect(response.statusCode).eql(200);
    for (var i = 0; i < sentEmail.attachments.length; i++) {
        await t.expect(sentEmail.attachments[i].name).eql(message.attachments[i].name);
    }
    logger.log('Verify sent data');

    //Get User2 mailbox
    var response = await user2.getMail();
    await t.expect(response.statusCode).eql(200);

    //Get mail from inbox matching the sent email
    let receivedEmail = user2.emails.find(e => e.emailId == sentEmail.emailId);
    await t.expect(receivedEmail).notEql(null);

    //Verify that the attachment data is the same as the sent attachments data
    var response = await user2.getAttachments(receivedEmail.id);
    await t
        .expect(response.statusCode).eql(200)
        .expect(receivedEmail.attachments.length).eql(sentEmail.attachments.length);
    logger.log('Load attachments');

    for (var i = 0; i < receivedEmail.attachments.length; i++) {
        await t
            .expect(receivedEmail.attachments[i].name).eql(sentEmail.attachments[i].name)
            .expect(receivedEmail.attachments[i].size).eql(sentEmail.attachments[i].size)
            .expect(receivedEmail.attachments[i].contentType).eql(sentEmail.attachments[i].contentType)
            .expect(receivedEmail.attachments[i].contentBytes).eql(sentEmail.attachments[i].contentBytes);
    }
    logger.log('Sent email attachments matches received email attachments');
    logger.log('-----PASSED-----');
})