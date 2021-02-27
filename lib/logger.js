export class Logger {
    constructor(testName) {
        this.testName = testName;
    }

    log(message) {
        console.log(this.testName + ': ' + message);
    }
}