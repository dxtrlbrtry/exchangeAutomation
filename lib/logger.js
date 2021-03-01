export default {
    init(context) {
        this.context = context;
    },
    log(message) {
        console.log('INFO: ' + this.context + ': ' + message);
    },
    error(message) {
        console.error('ERROR: ' + this.context + ': ' + message.toString());
    }
}