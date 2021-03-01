var Validator = require('jsonschema').Validator;
export const validator = new Validator();
export const schemas = {
    definitions: {
        authentication: {
            "type": "object",
            "properties": {
                "access_token": { "type": "string" },
                "token_type": { "type": "string" },
                "scope": { "type": "string" },
                "expires_in": { "type": "integer", "minimum": 1 }
            },
            "required": [ "access_token", "token_type", "scope", "expires_in" ]
        },
        user: {
            "type": "object",
            "properties": {
                "emailAddress": { 
                    "type": "object",
                    "properties": {
                        "name": "string",
                        "address": "string"
                    },
                    "required": [ "name", "address" ]
                }
            },
            "required": [ "emailAddress" ]
        },
        getMail: {
            "type": "object",
            "properties": {
                "@odata.context": { "type": "string" },
                "value": { 
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "@odata.etag": { "type": "string" },
                            "id": { "type": "string"},
                            "subject": { "type": "string" },
                            "body": {
                                "type": "object",
                                "properties": {
                                    "content": { "type": "string" }
                                },
                                "required": ["content"]
                            },
                            "internetMessageId": { "type": "string" },
                            "sender": { "$ref": "#/definitions/user" },
                            "from": { "$ref": "#/definitions/user" },
                            "toRecipients": {
                                "type": "array",
                                "items": { "$ref": "#/definitions/user" }
                            }
                        },
                        "required": [
                            "@odata.etag", "id", "subject", "body", 
                            "internetMessageId", "sender", "from", "toRecipients"
                        ]
                    }
                }
            },
            "required": ["@odata.context", "value"]
        },
        getAttachments: {
            "type": "object",
            "properties": {
                "name": { "type": "string" },
                "size": { "type": "integer", "minimum": 1 },
                "contentType": { "type": "string" },
                "contentBytes": { "type": "string" }
            },
            "required": [ "name", "size", "contentType", "contentBytes" ]
        }
    }
}