const createTestCafe = require('testcafe');

(async () => {
    const testcafe = await createTestCafe()
    try {
        await testcafe.createRunner()
            .src('tests/')
            .browsers(['chrome:headless'])
            .reporter('json')
            .run();
    }
    finally {
        await testcafe.close();
    }
})();