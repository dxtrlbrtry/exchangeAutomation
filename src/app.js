const createTestCafe = require('testcafe');

(async () => {
    const testCafe = await createTestCafe();
    try {
        const runner = testCafe.createRunner();
        await runner.run();
    }
    finally {
        testCafe.close();
    }
})();