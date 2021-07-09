"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
// This class is required when associating and disassociating since this is not supported by the Xrm.WebApi client side api at this time
var WebApiRequest;
(function (WebApiRequest) {
    var webApiUrl = "";
    function getWebApiUrl() {
        var context;
        var clientUrl;
        var apiVersion;
        if (webApiUrl)
            return webApiUrl;
        if (GetGlobalContext) {
            context = GetGlobalContext();
        }
        else {
            if (Xrm) {
                context = Xrm.Page.context;
            }
            else {
                throw new Error("Context is not available.");
            }
        }
        clientUrl = context.getClientUrl();
        var versionParts = context.getVersion().split(".");
        webApiUrl = clientUrl + "/api/data/v" + versionParts[0] + "." + versionParts[1];
        // Add the WebApi version
        return webApiUrl;
    }
    WebApiRequest.getWebApiUrl = getWebApiUrl;
    function getOdataContext() {
        return WebApiRequest.getWebApiUrl() + "/$metadata#$ref";
    }
    WebApiRequest.getOdataContext = getOdataContext;
    function request(action, uri, payload, includeFormattedValues, maxPageSize) {
        // Construct a fully qualified URI if a relative URI is passed in.
        if (uri.charAt(0) === "/") {
            uri = WebApiRequest.getWebApiUrl() + uri;
        }
        return new Promise(function (resolve, reject) {
            var request = new XMLHttpRequest();
            request.open(action, encodeURI(uri), true);
            request.setRequestHeader("OData-MaxVersion", "4.0");
            request.setRequestHeader("OData-Version", "4.0");
            request.setRequestHeader("Accept", "application/json");
            request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            if (maxPageSize) {
                request.setRequestHeader("Prefer", "odata.maxpagesize=" + maxPageSize);
            }
            if (includeFormattedValues) {
                request.setRequestHeader("Prefer", "odata.include-annotations=OData.Community.Display.V1.FormattedValue");
            }
            request.onreadystatechange = function () {
                if (this.readyState === 4) {
                    request.onreadystatechange = null;
                    switch (this.status) {
                        case 200: // Success with content returned in response body.
                        case 204: // Success with no content returned in response body.
                            resolve(this);
                            break;
                        default:
                            // All other statuses are unexpected so are treated like errors.
                            var error;
                            try {
                                error = JSON.parse(request.response).error;
                            }
                            catch (e) {
                                error = new Error("Unexpected Error");
                            }
                            reject(error);
                            break;
                    }
                }
            };
            request.send(JSON.stringify(payload));
        });
    }
    WebApiRequest.request = request;
})(WebApiRequest || (WebApiRequest = {}));
// This is required due to a bug where entity metadata is not loaded correctly in webresources that open in a new window
// Not needed if the webresource is embdedded inside a form
var windowStatic = window;
windowStatic.ENTITY_SET_NAMES = "{\n    \"account\": \"accounts\",\n    \"contact\": \"contacts\",\n    \"opportunity\": \"opportunities\"\n}";
windowStatic.ENTITY_PRIMARY_KEYS = "{\n    \"account\" : \"accountid\",\n    \"contact\" : \"contactid\",\n    \"opportunity\" : \"opportunityid\"\n}";
// Demonstrates the following techniques:
//  Using the Bound function AddToQueueResponse
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-actions#bound-actions
describe("", function () {
    it("AddToQueueResponse", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, queueid, letterid, AddToQueueRequest, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        return [4 /*yield*/, Xrm.WebApi.createRecord("queue", { "name": "Sample Queue" })];
                    case 1:
                        queueid = (_a.sent()).id;
                        return [4 /*yield*/, Xrm.WebApi.createRecord("letter", { "subject": "Sample Letter" })];
                    case 2:
                        letterid = (_a.sent()).id;
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 6, 11]);
                        AddToQueueRequest = new /** @class */ (function () {
                            function class_1() {
                                this.entity = {
                                    id: queueid,
                                    entityType: "queue"
                                };
                                this.Target = {
                                    id: letterid,
                                    entityType: "letter"
                                };
                            }
                            class_1.prototype.getMetadata = function () {
                                return {
                                    boundParameter: "entity",
                                    parameterTypes: {
                                        "entity": {
                                            typeName: "mscrm.queue",
                                            structuralProperty: 5
                                        },
                                        "QueueItemProperties": {
                                            typeName: "mscrm.queueitem",
                                            structuralProperty: 5
                                        },
                                        "SourceQueue": {
                                            typeName: "mscrm.queue",
                                            structuralProperty: 5
                                        },
                                        "Target": {
                                            typeName: "mscrm.crmbaseentity",
                                            structuralProperty: 5
                                        },
                                    },
                                    operationType: 0,
                                    operationName: "AddToQueue"
                                };
                            };
                            return class_1;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(AddToQueueRequest)];
                    case 4: return [4 /*yield*/, (_a.sent()).json()];
                    case 5:
                        response = _a.sent();
                        assert.isString(response.QueueItemId, "QueueItemId returned");
                        return [3 /*break*/, 11];
                    case 6:
                        if (!letterid) return [3 /*break*/, 8];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("letter", letterid)];
                    case 7:
                        _a.sent();
                        _a.label = 8;
                    case 8:
                        if (!queueid) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("queue", queueid)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following techniques:
//  1. Creating an opportunity
//  2. Winning an opportunity using the execute method
// See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions
describe("", function () {
    it("Win Opportunity", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, account1, createAccountResponse, opportunity1, createOpportunityResponse, winOpportunityRequest, rawResponse, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        account1 = {
                            name: "Sample Account"
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", account1)];
                    case 1:
                        createAccountResponse = _a.sent();
                        account1.accountid = createAccountResponse.id;
                        opportunity1 = {
                            name: "Sample Opportunity",
                            estimatedvalue: 1000,
                            estimatedclosedate: "2019-02-10",
                            "parentaccountid@odata.bind": "accounts(" + account1.accountid + ")"
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("opportunity", opportunity1)];
                    case 2:
                        createOpportunityResponse = _a.sent();
                        opportunity1.opportunityid = createOpportunityResponse.id;
                        winOpportunityRequest = new /** @class */ (function () {
                            function class_2() {
                                this.OpportunityClose = {
                                    description: "Sample Opportunity Close",
                                    subject: "Sample",
                                    "@odata.type": "Microsoft.Dynamics.CRM.opportunityclose",
                                    "opportunityid@odata.bind": "opportunities(" + opportunity1.opportunityid + ")"
                                };
                                this.Status = 3;
                            }
                            class_2.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {
                                        OpportunityClose: {
                                            typeName: "mscrm.opportunityclose",
                                            structuralProperty: 5
                                        },
                                        Status: {
                                            typeName: "Edm.Int32",
                                            structuralProperty: 1
                                        }
                                    },
                                    operationType: 0,
                                    operationName: "WinOpportunity"
                                };
                            };
                            return class_2;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(winOpportunityRequest)];
                    case 3:
                        rawResponse = (_a.sent());
                        return [4 /*yield*/, rawResponse.text()];
                    case 4:
                        response = _a.sent();
                        if (!response.id) return [3 /*break*/, 6];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", response.id)];
                    case 5:
                        _a.sent();
                        _a.label = 6;
                    case 6: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Creating a record
describe("", function () {
    it("Create", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, context, account1, _a, account2;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        context = GetGlobalContext();
                        account1 = {
                            name: "Sample Account",
                            accountcategorycode: 1,
                            creditlimit: 1000,
                            creditonhold: true,
                            numberofemployees: 10,
                            lastonholdtime: new Date(),
                            "preferredsystemuserid@odata.bind": "systemusers(" + context
                                .getUserId()
                                .replace("{", "")
                                .replace("}", "") + ")"
                        };
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, , 4, 7]);
                        // Create Account
                        _a = account1;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", account1))];
                    case 2:
                        // Create Account
                        _a.accountid = (_b.sent()).id;
                        if (!account1.accountid) {
                            throw new Error("Account not created");
                        }
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", account1.accountid, "?$select=name")];
                    case 3:
                        account2 = _b.sent();
                        assert.equal(account2.name, "Sample Account", "Account created");
                        return [3 /*break*/, 7];
                    case 4:
                        if (!account1.accountid) return [3 /*break*/, 6];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", account1.accountid)];
                    case 5:
                        _b.sent();
                        _b.label = 6;
                    case 6: return [7 /*endfinally*/];
                    case 7: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Creating an account and 2 related contacts in the same transaction
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api#associate-entities-on-create
describe("", function () {
    it("Deep Insert", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, context, account, _a, accountCreated;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        context = GetGlobalContext();
                        account = {
                            name: "Sample Account",
                            contact_customer_accounts: [
                                {
                                    firstname: "Sample",
                                    lastname: "contact 1"
                                },
                                {
                                    firstname: "Sample",
                                    lastname: "Contact 2"
                                },
                            ]
                        };
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, , 4, 7]);
                        // Create Account & Contacts
                        _a = account;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", account))];
                    case 2:
                        // Create Account & Contacts
                        _a.accountid = (_b.sent()).id;
                        if (!account.accountid)
                            throw new Error("Account not created");
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", account.accountid, "?$select=name&$expand=contact_customer_accounts($select=firstname,lastname)")];
                    case 3:
                        accountCreated = _b.sent();
                        assert.equal(accountCreated.contact_customer_accounts.length, 2, "Account created with 2 contacts");
                        return [3 /*break*/, 7];
                    case 4:
                        if (!account.accountid) return [3 /*break*/, 6];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", account.accountid)];
                    case 5:
                        _b.sent();
                        _b.label = 6;
                    case 6: return [7 /*endfinally*/];
                    case 7: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Deleting a record
describe("", function () {
    it("Delete", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, record, _a, fetch, accounts;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        record = {
                            name: "Sample Account"
                        };
                        // Create Account
                        _a = record;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", record))];
                    case 1:
                        // Create Account
                        _a.accountid = (_b.sent()).id;
                        if (!record.accountid) return [3 /*break*/, 3];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", record.accountid)];
                    case 2:
                        _b.sent();
                        _b.label = 3;
                    case 3:
                        fetch = "<fetch no-lock=\"true\" >\n       <entity name=\"account\" >\n         <filter>\n           <condition attribute=\"accountid\" operator=\"eq\" value=\"" + record.accountid + "\" />\n         </filter>\n       </entity>\n     </fetch>";
                        return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("account", "?fetchXml=" + fetch)];
                    case 4:
                        accounts = _b.sent();
                        assert.equal(accounts.entities.length, 0, "Account deleted");
                        return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Creating a record and then retrieving it unchanged with the same etag
//  This is a technique employed by the client side api to avoid uneccesary data being transferred over the network
//  This sample shows how varying the options will invalidate the client side cache
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api
describe("", function () {
    it("etags with $expand", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, contactid, accountid, query1, etag1, query2, etag2, query3, etag3, query4, etag4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact"
                            })];
                    case 1:
                        contactid = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account",
                                "primarycontactid@odata.bind": "/contacts(" + contactid.id + ")"
                            })];
                    case 2:
                        accountid = _a.sent();
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 9, 14]);
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", accountid.id, "?$select=name&$expand=primarycontactid")];
                    case 4:
                        query1 = _a.sent();
                        etag1 = query1["@odata.etag"];
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", accountid.id, "?$select=name&$expand=primarycontactid")];
                    case 5:
                        query2 = _a.sent();
                        etag2 = query2["@odata.etag"];
                        assert.equal(etag1, etag2, "Record not modified");
                        assert.equal("Sample Contact", query2.primarycontactid.lastname, "Related contact returned");
                        // Update the contact name
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("contact", contactid.id, {
                                lastname: "Sample Contact (edited)"
                            })];
                    case 6:
                        // Update the contact name
                        _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", accountid.id, "?$select=name&$expand=primarycontactid")];
                    case 7:
                        query3 = _a.sent();
                        etag3 = query3["@odata.etag"];
                        assert.equal(etag1, etag3, "Record not modified");
                        assert.equal("Sample Contact", query3.primarycontactid.lastname, "Unchanged contact");
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", accountid.id, "?$select=name&$expand=primarycontactid($select=lastname)")];
                    case 8:
                        query4 = _a.sent();
                        etag4 = query4["@odata.etag"];
                        assert.equal(etag1, etag4, "Record not modified");
                        assert.equal("Sample Contact (edited)", query4.primarycontactid.lastname, "Contact changed");
                        return [3 /*break*/, 14];
                    case 9:
                        if (!accountid.id) return [3 /*break*/, 11];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 10:
                        _a.sent();
                        _a.label = 11;
                    case 11:
                        if (!contactid.id) return [3 /*break*/, 13];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contactid.id)];
                    case 12:
                        _a.sent();
                        _a.label = 13;
                    case 13: return [7 /*endfinally*/];
                    case 14: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Creating a record and then retrieving it unchanged with the same etag
//  This is a technique employed by the client side api to avoid uneccesary data being transferred over the network
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api
describe("", function () {
    it("etags", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, account, _a, query1, etag1, query2, etag2, query3, etag3;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        account = {
                            name: "Sample Account",
                            creditlimit: 1000,
                        };
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, , 7, 10]);
                        // Create Account
                        _a = account;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", account))];
                    case 2:
                        // Create Account
                        _a.accountid = (_b.sent()).id;
                        if (!account.accountid)
                            throw new Error("Account not created");
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", account.accountid, "?$select=name")];
                    case 3:
                        query1 = _b.sent();
                        etag1 = query1["@odata.etag"];
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", account.accountid, "?$select=name")];
                    case 4:
                        query2 = _b.sent();
                        etag2 = query2["@odata.etag"];
                        assert.equal(etag1, etag2, "Record not modified");
                        // Update the value
                        account.name = "Sample Account (updated)";
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("account", account.accountid, account)];
                    case 5:
                        _b.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", account.accountid, "?$select=name")];
                    case 6:
                        query3 = _b.sent();
                        etag3 = query3["@odata.etag"];
                        assert.notEqual(etag1, etag3, "Record modified");
                        return [3 /*break*/, 10];
                    case 7:
                        if (!account.accountid) return [3 /*break*/, 9];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", account.accountid)];
                    case 8:
                        _b.sent();
                        _b.label = 9;
                    case 9: return [7 /*endfinally*/];
                    case 10: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Querying records using fetchxml
describe("", function () {
    it("Query with FetchXml", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, fetch, accounts;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        fetch = "<fetch no-lock=\"true\" >\n      <entity name=\"account\">\n        <attribute name=\"name\"/>\n      </entity>\n    </fetch>";
                        return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("account", "?fetchXml=" + fetch)];
                    case 1:
                        accounts = _a.sent();
                        assert.isNotNull(accounts.entities, "Account query returns results");
                        return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Querying for a record by id using retrieveRecord with a $select clause
describe("", function () {
    it("Read", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, record, _a, accountsRead;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        record = {
                            name: "Sample Account"
                        };
                        // Create Account
                        _a = record;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", record))];
                    case 1:
                        // Create Account
                        _a.accountid = (_b.sent()).id;
                        if (!record.accountid)
                            throw new Error("Account not created");
                        _b.label = 2;
                    case 2:
                        _b.trys.push([2, , 4, 7]);
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", record.accountid, "?$select=name,primarycontactid")];
                    case 3:
                        accountsRead = _b.sent();
                        if (!accountsRead || !accountsRead.name) {
                            throw new Error("Account not created");
                        }
                        assert.equal(accountsRead.name, record.name, "Account created");
                        return [3 /*break*/, 7];
                    case 4:
                        if (!record.accountid) return [3 /*break*/, 6];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", record.accountid)];
                    case 5:
                        _b.sent();
                        _b.label = 6;
                    case 6: return [7 /*endfinally*/];
                    case 7: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Querying for multiple records using odata query
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/query-data-web-api
describe("", function () {
    it("RetrieveMultiple", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, accountid, results;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assert = chai.assert;
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account",
                                revenue: 20000.01
                            })];
                    case 1:
                        accountid = _a.sent();
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, , 4, 7]);
                        return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("account", "?$select=name&$filter=revenue gt 20000 and revenue lt 20001 and name eq 'Sample Account'", 10)];
                    case 3:
                        results = _a.sent();
                        // Check that there is a single result returned
                        if (!results.entities || !results.entities.length)
                            throw new Error("No results returned");
                        assert.equal(results.entities.length, 1, "Single result returned");
                        return [3 /*break*/, 7];
                    case 4:
                        if (!accountid.id) return [3 /*break*/, 6];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 5:
                        _a.sent();
                        _a.label = 6;
                    case 6: return [7 /*endfinally*/];
                    case 7: return [2 /*return*/];
                }
            });
        });
    });
});
describe("", function () {
    it("SetState", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, account, _a, opportunity, createOpportunityResponse, opportunityRead, accountRead;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        account = {
                            name: "Sample Account"
                        };
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, , 8, 11]);
                        // Create Account
                        _a = account;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", account))];
                    case 2:
                        // Create Account
                        _a.accountid = (_b.sent()).id;
                        if (!account.accountid) {
                            throw new Error("Account ID not returned");
                        }
                        opportunity = {
                            name: "Sample Opportunity",
                            "parentaccountid@odata.bind": "accounts(" + account.accountid + ")"
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("opportunity", opportunity)];
                    case 3:
                        createOpportunityResponse = _b.sent();
                        opportunity.opportunityid = createOpportunityResponse.id;
                        // Change Opportunity Status Reason to In Progress
                        opportunity.statuscode = 1 /* In_Progress */;
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("opportunity", opportunity.opportunityid, opportunity)];
                    case 4:
                        _b.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("opportunity", opportunity.opportunityid, "?$select=statuscode")];
                    case 5:
                        opportunityRead = _b.sent();
                        if (!opportunityRead || !opportunityRead.statuscode) {
                            throw new Error("Opportunity not updated");
                        }
                        assert.equal(opportunityRead.statuscode, 1 /* In_Progress */, "Opportunity In Progress");
                        // Update account state to In Active
                        account.statecode = 1 /* Inactive */;
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("account", account.accountid, account)];
                    case 6:
                        _b.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", account.accountid, "?$select=statecode")];
                    case 7:
                        accountRead = _b.sent();
                        if (!accountRead || accountRead.statecode == undefined) {
                            throw new Error("Account not updated");
                        }
                        assert.equal(accountRead.statecode, 1 /* Inactive */, "Account Inactive");
                        return [3 /*break*/, 11];
                    case 8:
                        if (!account.accountid) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", account.accountid)];
                    case 9:
                        _b.sent();
                        _b.label = 10;
                    case 10: return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  Updating a record
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api
describe("", function () {
    it("Update", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, record, _a, accountsRead;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        record = {
                            name: "Sample Account"
                        };
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, , 5, 8]);
                        // Create Account
                        _a = record;
                        return [4 /*yield*/, (Xrm.WebApi.createRecord("account", record))];
                    case 2:
                        // Create Account
                        _a.accountid = (_b.sent()).id;
                        if (!record.accountid) {
                            throw new Error("Account ID not returned");
                        }
                        record.name = "Sample Account (updated)";
                        record.address1_city = "Oxford";
                        // Update account
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("account", record.accountid, record)];
                    case 3:
                        // Update account
                        _b.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("account", record.accountid, "?$select=name,address1_city")];
                    case 4:
                        accountsRead = _b.sent();
                        if (!accountsRead || !accountsRead.name || !accountsRead.address1_city) {
                            throw new Error("Account not updated");
                        }
                        assert.equal(accountsRead.name, record.name, "Account updated");
                        assert.equal(accountsRead.address1_city, record.address1_city, "Account updated");
                        return [3 /*break*/, 8];
                    case 5:
                        if (!record.accountid) return [3 /*break*/, 7];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", record.accountid)];
                    case 6:
                        _b.sent();
                        _b.label = 7;
                    case 7: return [7 /*endfinally*/];
                    case 8: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following techniques:
//  Using the CalculateRollup Unbound Function to recalculate a rollup field
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions
describe("", function () {
    it("CalculateRollup", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, response, request;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("account", "?$select=accountid,name&$top=1", 1)];
                    case 1:
                        response = _a.sent();
                        request = new /** @class */ (function () {
                            function class_3() {
                                this.Target = {
                                    id: response.entities[0].accountid,
                                    entityType: "account"
                                };
                                this.FieldName = "opendeals";
                            }
                            class_3.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {
                                        FieldName: {
                                            typeName: "Edm.String",
                                            structuralProperty: 1
                                        },
                                        Target: {
                                            typeName: "mscrm.crmbaseentity",
                                            structuralProperty: 5
                                        }
                                    },
                                    operationType: 1,
                                    operationName: "CalculateRollupField"
                                };
                            };
                            return class_3;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(request)];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following techniques:
//  Using the CalculateTotalTimeIncident Bound Function to recalculate a rollup field
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions
describe("", function () {
    it("CalculateTotalTimeIncident", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, response, request, rawResponse;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("incident", "?$select=incidentid", 1)];
                    case 1:
                        response = _a.sent();
                        request = new /** @class */ (function () {
                            function class_4() {
                                this.entity = {
                                    id: response.entities[0].incidentid,
                                    entityType: "incident"
                                };
                            }
                            class_4.prototype.getMetadata = function () {
                                return {
                                    boundParameter: "entity",
                                    parameterTypes: {
                                        entity: {
                                            typeName: "mscrm.incident",
                                            structuralProperty: 5
                                        }
                                    },
                                    operationType: 1,
                                    operationName: "CalculateTotalTimeIncident"
                                };
                            };
                            return class_4;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(request)];
                    case 2: return [4 /*yield*/, (_a.sent()).json()];
                    case 3:
                        rawResponse = _a.sent();
                        ;
                        assert.isNumber(rawResponse.TotalTime, "Total Time returned");
                        return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following techniques:
//  Querying metadata using the RetrieveMetadataChanges request
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions
describe("", function () {
    it("RetrieveMetadataChanges", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, request, rawResponse, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        request = new /** @class */ (function () {
                            function class_5() {
                                this.Query = {
                                    Criteria: {
                                        Conditions: [
                                            {
                                                PropertyName: "LogicalName",
                                                ConditionOperator: "Equals",
                                                Value: {
                                                    Value: "account",
                                                    Type: "System.String"
                                                }
                                            }
                                        ],
                                        FilterOperator: "And"
                                    },
                                    Properties: {
                                        PropertyNames: ["Attributes"]
                                    },
                                    AttributeQuery: {
                                        Properties: {
                                            PropertyNames: ["OptionSet"]
                                        },
                                        Criteria: {
                                            Conditions: [
                                                {
                                                    PropertyName: "LogicalName",
                                                    ConditionOperator: "Equals",
                                                    Value: {
                                                        Value: "address1_shippingmethodcode",
                                                        Type: "System.String"
                                                    }
                                                }
                                            ],
                                            FilterOperator: "And"
                                        }
                                    }
                                };
                            }
                            class_5.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {
                                        AppModuleId: {
                                            typeName: "Edm.Guid",
                                            structuralProperty: 1
                                        },
                                        ClientVersionStamp: {
                                            typeName: "Edm.String",
                                            structuralProperty: 1
                                        },
                                        DeletedMetadataFilters: {
                                            typeName: "mscrm.DeletedMetadataFilters",
                                            structuralProperty: 3
                                        },
                                        Query: {
                                            typeName: "mscrm.EntityQueryExpression",
                                            structuralProperty: 5
                                        }
                                    },
                                    operationType: 1,
                                    operationName: "RetrieveMetadataChanges"
                                };
                            };
                            return class_5;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(request)];
                    case 1:
                        rawResponse = _a.sent();
                        return [4 /*yield*/, rawResponse.json()];
                    case 2:
                        response = _a.sent();
                        assert.equal(response.EntityMetadata[0].LogicalName, "account", "Account metadata returned");
                        assert.ok(response.EntityMetadata[0].Attributes.length > 0, "Account Attributes returned");
                        return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following techniques:
//  1. Creating a record with a lookup field
//  2. Updating a lookup field on an existing record
//  3. Setting a lookup field to null on an existing record
//  4. Querying Lookups
/// <reference path="../WebApiRequest.ts" />
describe("", function () {
    it("Lookup Fields", function () {
        return __awaiter(this, void 0, void 0, function () {
            var contact1id, contact2id, accountid, url, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact 1"
                            })];
                    case 1:
                        contact1id = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact 2"
                            })];
                    case 2:
                        contact2id = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account",
                                "primarycontactid@odata.bind": "/contacts(" + contact1id.id + ")"
                            })];
                    case 3:
                        accountid = _a.sent();
                        _a.label = 4;
                    case 4:
                        _a.trys.push([4, , 7, 14]);
                        // Update the primary contact to a different contact
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("account", accountid.id, {
                                "primarycontactid@odata.bind": "/contacts(" + contact2id.id + ")"
                            })];
                    case 5:
                        // Update the primary contact to a different contact
                        _a.sent();
                        url = "/accounts(" + accountid.id + ")/primarycontactid/$ref";
                        return [4 /*yield*/, WebApiRequest.request("DELETE", url)];
                    case 6:
                        response = _a.sent();
                        return [3 /*break*/, 14];
                    case 7:
                        if (!accountid) return [3 /*break*/, 9];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 8:
                        _a.sent();
                        _a.label = 9;
                    case 9:
                        if (!contact1id) return [3 /*break*/, 11];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contact1id.id)];
                    case 10:
                        _a.sent();
                        _a.label = 11;
                    case 11:
                        if (!contact2id) return [3 /*break*/, 13];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contact2id.id)];
                    case 12:
                        _a.sent();
                        _a.label = 13;
                    case 13: return [7 /*endfinally*/];
                    case 14: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  1. Associating two records over a many to many relationship
//  2. Disassociating the two records
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api
describe("", function () {
    it("Many to Many", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, accountid, leadid, associate, url, response, fetch, associatedLeads, url, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        assert = chai.assert;
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account"
                            })];
                    case 1:
                        accountid = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("lead", {
                                lastname: "Sample Lead"
                            })];
                    case 2:
                        leadid = _a.sent();
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 7, 12]);
                        associate = {
                            "@odata.context": WebApiRequest.getOdataContext(),
                            "@odata.id": "leads(" + leadid.id + ")"
                        };
                        url = "/accounts(" + accountid.id + ")/accountleads_association/$ref";
                        return [4 /*yield*/, WebApiRequest.request("POST", url, associate)];
                    case 4:
                        response = _a.sent();
                        fetch = "<fetch no-lock=\"true\" >\n      <entity name=\"account\" >\n        <attribute name=\"name\" />\n        <filter>\n          <condition attribute=\"accountid\" operator=\"eq\" value=\"" + accountid.id + "\" />\n        </filter>\n        <link-entity name=\"accountleads\" from=\"accountid\" to=\"accountid\" intersect=\"true\" >\n          <attribute name=\"name\" />\n        </link-entity>\n      </entity>\n    </fetch>";
                        return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords("account", "?fetchXml=" + fetch)];
                    case 5:
                        associatedLeads = _a.sent();
                        assert.equal(associatedLeads.entities.length, 1, "Associated records");
                        url = "/accounts(" + accountid.id + ")/accountleads_association(" + leadid.id + ")/$ref";
                        return [4 /*yield*/, WebApiRequest.request("DELETE", url)];
                    case 6:
                        response = _a.sent();
                        return [3 /*break*/, 12];
                    case 7:
                        if (!leadid) return [3 /*break*/, 9];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("lead", leadid.id)];
                    case 8:
                        _a.sent();
                        _a.label = 9;
                    case 9:
                        if (!accountid) return [3 /*break*/, 11];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 10:
                        _a.sent();
                        _a.label = 11;
                    case 11: return [7 /*endfinally*/];
                    case 12: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  1. Associating two records over a many-to-many relationship
//  2. Disassociating the two records
//  NOTE: This sample uses the execute method with a operationName of 'Associate' and 'Disassociate'
describe("", function () {
    it("Many to One using execute", function () {
        return __awaiter(this, void 0, void 0, function () {
            var contactid, accountid, associateRequest, response, dissassociateRequest, dissassociateResponse;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact"
                            })];
                    case 1:
                        contactid = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account"
                            })];
                    case 2:
                        accountid = _a.sent();
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 6, 11]);
                        associateRequest = new /** @class */ (function () {
                            function class_6() {
                                this.target = {
                                    id: accountid.id,
                                    entityType: "account"
                                };
                                this.relatedEntities = [
                                    {
                                        id: contactid.id,
                                        entityType: "contact"
                                    }
                                ];
                                this.relationship = "account_primary_contact";
                            }
                            class_6.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {},
                                    operationType: 2,
                                    operationName: "Associate"
                                };
                            };
                            return class_6;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(associateRequest)];
                    case 4:
                        response = _a.sent();
                        dissassociateRequest = new /** @class */ (function () {
                            function class_7() {
                                this.target = {
                                    id: accountid.id,
                                    entityType: "account"
                                };
                                this.relationship = "primarycontactid";
                            }
                            class_7.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {},
                                    operationType: 2,
                                    operationName: "Disassociate"
                                };
                            };
                            return class_7;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(dissassociateRequest)];
                    case 5:
                        dissassociateResponse = _a.sent();
                        return [3 /*break*/, 11];
                    case 6:
                        if (!accountid) return [3 /*break*/, 8];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 7:
                        _a.sent();
                        _a.label = 8;
                    case 8:
                        if (!contactid) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contactid.id)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
/// <reference path="../WebApiRequest.ts" />
// Demonstrates the following technique:
//  1. Associating two records over a many-to-one relationship. This is an alternative to simply using a lookup field in a create/update
//  2. Disassociating the two records
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api
describe("", function () {
    it("Many to One", function () {
        return __awaiter(this, void 0, void 0, function () {
            var contactid, accountid, associate, url, response, url, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact"
                            })];
                    case 1:
                        contactid = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account"
                            })];
                    case 2:
                        accountid = _a.sent();
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 6, 11]);
                        associate = {
                            "@odata.context": WebApiRequest.getOdataContext(),
                            "@odata.id": "contacts(" + contactid.id + ")"
                        };
                        url = "/accounts(" + accountid.id + ")/primarycontactid/$ref";
                        return [4 /*yield*/, WebApiRequest.request("PUT", url, associate)];
                    case 4:
                        response = _a.sent();
                        url = "/accounts(" + accountid.id + ")/primarycontactid/$ref";
                        return [4 /*yield*/, WebApiRequest.request("DELETE", url)];
                    case 5:
                        response = _a.sent();
                        return [3 /*break*/, 11];
                    case 6:
                        if (!accountid) return [3 /*break*/, 8];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 7:
                        _a.sent();
                        _a.label = 8;
                    case 8:
                        if (!contactid) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contactid.id)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following technique:
//  1. Associating two records over a one-to-many relationship
//  2. Disassociating the two records
//  NOTE: This sample uses the execute method with a operationName of 'Associate' and 'Disassociate'
describe("", function () {
    it("One to Many using execute", function () {
        return __awaiter(this, void 0, void 0, function () {
            var contactid, accountid, associateRequest, response, dissassociateRequest, dissassociateResponse;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact"
                            })];
                    case 1:
                        contactid = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account"
                            })];
                    case 2:
                        accountid = _a.sent();
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 6, 11]);
                        associateRequest = new /** @class */ (function () {
                            function class_8() {
                                this.target = {
                                    id: contactid.id,
                                    entityType: "contact"
                                };
                                this.relatedEntities = [
                                    {
                                        id: accountid.id,
                                        entityType: "account"
                                    }
                                ];
                                this.relationship = "account_primary_contact";
                            }
                            class_8.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {},
                                    operationType: 2,
                                    operationName: "Associate"
                                };
                            };
                            return class_8;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(associateRequest)];
                    case 4:
                        response = _a.sent();
                        dissassociateRequest = new /** @class */ (function () {
                            function class_9() {
                                this.target = {
                                    id: contactid.id,
                                    entityType: "contact"
                                };
                                this.relatedEntityId = accountid.id;
                                this.relationship = "account_primary_contact";
                            }
                            class_9.prototype.getMetadata = function () {
                                return {
                                    parameterTypes: {},
                                    operationType: 2,
                                    operationName: "Disassociate"
                                };
                            };
                            return class_9;
                        }())();
                        return [4 /*yield*/, Xrm.WebApi.online.execute(dissassociateRequest)];
                    case 5:
                        dissassociateResponse = _a.sent();
                        return [3 /*break*/, 11];
                    case 6:
                        if (!accountid) return [3 /*break*/, 8];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 7:
                        _a.sent();
                        _a.label = 8;
                    case 8:
                        if (!contactid) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contactid.id)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
/// <reference path="../WebApiRequest.ts" />
// Demonstrates the following technique:
//  1. Associating two records over a one-to-many relationship
//  2. Disassociating the two records
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api
describe("", function () {
    it("One to Many", function () {
        return __awaiter(this, void 0, void 0, function () {
            var contactid, accountid, associate, url, response, url, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", {
                                lastname: "Sample Contact"
                            })];
                    case 1:
                        contactid = _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", {
                                name: "Sample Account"
                            })];
                    case 2:
                        accountid = _a.sent();
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 6, 11]);
                        associate = {
                            "@odata.context": WebApiRequest.getOdataContext(),
                            "@odata.id": "accounts(" + accountid.id + ")"
                        };
                        url = "/contacts(" + contactid.id + ")/account_primary_contact/$ref";
                        return [4 /*yield*/, WebApiRequest.request("POST", url, associate)];
                    case 4:
                        response = _a.sent();
                        url = "/contacts(" + contactid.id + ")/account_primary_contact(" + accountid.id + ")/$ref";
                        return [4 /*yield*/, WebApiRequest.request("DELETE", url)];
                    case 5:
                        response = _a.sent();
                        return [3 /*break*/, 11];
                    case 6:
                        if (!accountid) return [3 /*break*/, 8];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", accountid.id)];
                    case 7:
                        _a.sent();
                        _a.label = 8;
                    case 8:
                        if (!contactid) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contactid.id)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
// Demonstrates the following techniques:
//  1. Creating activities with activity parties
//  2. Updating activity parties - this is a special case
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api#associate-entities-on-update-using-collection-valued-navigation-property
/// <reference path="../WebApiRequest.ts" />
describe("", function () {
    it("Activity Parties Letter", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, contact1, contact1id, contact2, contact2id, letter1, letter1id, letter2, partyTo, letter3, partyTo2, partyTo3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(900000);
                        assert = chai.assert;
                        contact1 = {
                            lastname: "Test Contact 1" + new Date().toUTCString()
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", contact1)];
                    case 1:
                        contact1id = (_a.sent())
                            .id;
                        contact2 = {
                            lastname: "Test Contact 2 " + new Date().toUTCString()
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", contact2)];
                    case 2:
                        contact2id = (_a.sent())
                            .id;
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 8, 13]);
                        letter1 = {
                            subject: "Sample Letter " + new Date().toUTCString(),
                            letter_activity_parties: [
                                {
                                    participationtypemask: 2,
                                    "@odata.type": "Microsoft.Dynamics.CRM.activityparty",
                                    "partyid_contact@odata.bind": "contacts(" + contact1id + ")"
                                }
                            ]
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("letter", letter1)];
                    case 4:
                        letter1id = (_a.sent())
                            .id;
                        if (!letter1id)
                            throw new Error("Letter not created");
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("letter", letter1id, "?$expand=letter_activity_parties($select=activitypartyid,_partyid_value,participationtypemask)")];
                    case 5:
                        letter2 = _a.sent();
                        if (!letter2.letter_activity_parties ||
                            !letter1.letter_activity_parties.length)
                            throw new Error("Letter1 letter_activity_parties not set");
                        partyTo = findPartyById(letter2.letter_activity_parties, contact1id);
                        assert.isNotNull(partyTo, "To Party returned");
                        // Add an activity party
                        letter1.letter_activity_parties.push({
                            participationtypemask: 2,
                            "@odata.type": "Microsoft.Dynamics.CRM.activityparty",
                            "partyid_contact@odata.bind": "contacts(" + contact2id + ")"
                        });
                        // Update letter
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("letter", letter1id, letter1)];
                    case 6:
                        // Update letter
                        _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("letter", letter1id, "?$expand=letter_activity_parties($select=activitypartyid,_partyid_value,participationtypemask)")];
                    case 7:
                        letter3 = _a.sent();
                        partyTo2 = findPartyById(letter3.letter_activity_parties, contact1id);
                        partyTo3 = findPartyById(letter3.letter_activity_parties, contact2id);
                        assert.isNotNull(partyTo2, "To Party 1 returned");
                        assert.isNotNull(partyTo3, "To Party 2 returned");
                        return [3 /*break*/, 13];
                    case 8:
                        if (!contact1id) return [3 /*break*/, 10];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contact1id)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10:
                        if (!contact2id) return [3 /*break*/, 12];
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contact2id)];
                    case 11:
                        _a.sent();
                        _a.label = 12;
                    case 12: return [7 /*endfinally*/];
                    case 13: return [2 /*return*/];
                }
            });
        });
    });
    function findPartyById(parties, id) {
        for (var _i = 0, parties_1 = parties; _i < parties_1.length; _i++) {
            var party = parties_1[_i];
            if (party._partyid_value == id) {
                return party;
            }
        }
        return null;
    }
});
// Demonstrates the following techniques:
//  Creating records with customer fields that reference either an account or a contact
/// <reference path="../WebApiRequest.ts" />
describe("", function () {
    it("Customer Fields", function () {
        return __awaiter(this, void 0, void 0, function () {
            var assert, account1, account1id, contact1, contact1id, opportunity1, opportunity1id, opportunity2, opportunity3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.timeout(90000);
                        console.log("Creating account");
                        assert = chai.assert;
                        account1 = {
                            name: "Sample Account " + new Date().toUTCString()
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("account", account1)];
                    case 1:
                        account1id = (_a.sent())
                            .id;
                        if (!account1id)
                            throw new Error("accountid not defined");
                        contact1 = {
                            lastname: "Sample Contact " + new Date().toUTCString()
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("contact", contact1)];
                    case 2:
                        contact1id = (_a.sent())
                            .id;
                        if (!contact1id)
                            throw new Error("contact1id not defined");
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 8, 11]);
                        opportunity1 = {
                            name: "Sample Opportunity " + new Date().toUTCString(),
                            estimatedvalue: 1000,
                            estimatedclosedate: new Date(Date.now()).toISOString().substr(0, 10),
                            "customerid_account@odata.bind": "accounts(" + account1id + ")"
                        };
                        return [4 /*yield*/, Xrm.WebApi.createRecord("opportunity", opportunity1)];
                    case 4:
                        opportunity1id = (_a.sent()).id;
                        if (!opportunity1id)
                            throw new Error("Opportunty ID not defined");
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("opportunity", opportunity1id, "?$select=name,_customerid_value")];
                    case 5:
                        opportunity2 = _a.sent();
                        if (!opportunity2 || !opportunity2._customerid_value)
                            throw new Error("Opportunity2 Customerid not returned");
                        // Check that the customerid field is populated
                        assert.isNotEmpty(opportunity2._customerid_value, "Customer field not empty");
                        assert.equal(opportunity2._customerid_value, account1id, "Customer id equal to accountid");
                        assert.equal(opportunity2["_customerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"], "account", "Logical name set");
                        // Update the customer field to reference the contact
                        delete opportunity1["customerid_account@odata.bind"];
                        opportunity1["customerid_contact@odata.bind"] = "contacts(" + contact1id + ")";
                        return [4 /*yield*/, Xrm.WebApi.updateRecord("opportunity", opportunity1id, opportunity1)];
                    case 6:
                        _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord("opportunity", opportunity1id, "?$select=name,_customerid_value")];
                    case 7:
                        opportunity3 = _a.sent();
                        // Check that the customerid field is populated
                        assert.isNotEmpty(opportunity3._customerid_value, "Customer field not empty");
                        assert.equal(opportunity3._customerid_value, contact1id, "Customer id equal to contact");
                        assert.equal(opportunity3["_customerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"], "contact", "Logical name set");
                        return [3 /*break*/, 11];
                    case 8: 
                    // Delete the opportunity and account - opportunity is a cascade delete
                    return [4 /*yield*/, Xrm.WebApi.deleteRecord("contact", contact1id)];
                    case 9:
                        // Delete the opportunity and account - opportunity is a cascade delete
                        _a.sent();
                        return [4 /*yield*/, Xrm.WebApi.deleteRecord("account", account1id)];
                    case 10:
                        _a.sent();
                        return [7 /*endfinally*/];
                    case 11: return [2 /*return*/];
                }
            });
        });
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibW9kZWwtZHJpdmVuLXdlYmFwaS5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbInNyYy90ZXN0L1dlYkFwaVJlcXVlc3QudHMiLCJzcmMvdGVzdC9maXhYcm1XZWJBcGkudHMiLCJzcmMvdGVzdC9BY3Rpb25zL0FkZFRvUXVldWVSZXNwb25zZS50cyIsInNyYy90ZXN0L0FjdGlvbnMvV2luT3Bwb3J0dW5pdHkudHMiLCJzcmMvdGVzdC9DUlVEL0NyZWF0ZS50cyIsInNyYy90ZXN0L0NSVUQvRGVlcCBJbnNlcnQudHMiLCJzcmMvdGVzdC9DUlVEL0RlbGV0ZS50cyIsInNyYy90ZXN0L0NSVUQvRVRhZ3Mgd2l0aCBleHBhbmQudHMiLCJzcmMvdGVzdC9DUlVEL0VUYWdzLnRzIiwic3JjL3Rlc3QvQ1JVRC9GZXRjaFhtbC50cyIsInNyYy90ZXN0L0NSVUQvUmVhZC50cyIsInNyYy90ZXN0L0NSVUQvUmV0cmlldmVNdWx0aXBsZS50cyIsInNyYy90ZXN0L0NSVUQvU2V0U3RhdGUudHMiLCJzcmMvdGVzdC9DUlVEL1VwZGF0ZS50cyIsInNyYy90ZXN0L0Z1bmN0aW9ucy9DYWxjdWxhdGVSb2xsdXAudHMiLCJzcmMvdGVzdC9GdW5jdGlvbnMvQ2FsY3VsYXRlVG90YWxUaW1lSW5jaWRlbnQudHMiLCJzcmMvdGVzdC9GdW5jdGlvbnMvUmV0cmlldmVNZXRhZGF0YUNoYW5nZXMudHMiLCJzcmMvdGVzdC9SZWxhdGlvbnNoaXBzL0xvb2t1cCBGaWVsZHMudHMiLCJzcmMvdGVzdC9SZWxhdGlvbnNoaXBzL01hbnktdG8tTWFueS50cyIsInNyYy90ZXN0L1JlbGF0aW9uc2hpcHMvTWFueS10by1PbmUgdXNpbmcgZXhlY3V0ZS50cyIsInNyYy90ZXN0L1JlbGF0aW9uc2hpcHMvTWFueS10by1PbmUudHMiLCJzcmMvdGVzdC9SZWxhdGlvbnNoaXBzL09uZS10by1NYW55IHVzaW5nIGV4ZWN1dGUudHMiLCJzcmMvdGVzdC9SZWxhdGlvbnNoaXBzL09uZS10by1NYW55LnRzIiwic3JjL3Rlc3QvU3BlY2lhbC9BY3Rpdml0eVBhcnRpZXMudHMiLCJzcmMvdGVzdC9TcGVjaWFsL0N1c3RvbWVyRmllbGRzLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7OztBQUFBLHdJQUF3STtBQUN4SSxJQUFVLGFBQWEsQ0FzRnRCO0FBdEZELFdBQVUsYUFBYTtJQUNyQixJQUFJLFNBQVMsR0FBVyxFQUFFLENBQUM7SUFDM0IsU0FBZ0IsWUFBWTtRQUMxQixJQUFJLE9BQTBCLENBQUM7UUFDL0IsSUFBSSxTQUFpQixDQUFDO1FBQ3RCLElBQUksVUFBa0IsQ0FBQztRQUN2QixJQUFJLFNBQVM7WUFBRSxPQUFPLFNBQVMsQ0FBQztRQUVoQyxJQUFJLGdCQUFnQixFQUFFO1lBQ3BCLE9BQU8sR0FBRyxnQkFBZ0IsRUFBRSxDQUFDO1NBQzlCO2FBQU07WUFDTCxJQUFJLEdBQUcsRUFBRTtnQkFDUCxPQUFPLEdBQUcsR0FBRyxDQUFDLElBQUksQ0FBQyxPQUFPLENBQUM7YUFDNUI7aUJBQU07Z0JBQ0wsTUFBTSxJQUFJLEtBQUssQ0FBQywyQkFBMkIsQ0FBQyxDQUFDO2FBQzlDO1NBQ0Y7UUFDRCxTQUFTLEdBQUcsT0FBTyxDQUFDLFlBQVksRUFBRSxDQUFDO1FBQ25DLElBQUksWUFBWSxHQUFHLE9BQU8sQ0FBQyxVQUFVLEVBQUUsQ0FBQyxLQUFLLENBQUMsR0FBRyxDQUFDLENBQUM7UUFFbEQsU0FBUyxHQUFNLFNBQVMsbUJBQWMsWUFBWSxDQUFDLENBQUMsQ0FBQyxTQUNwRCxZQUFZLENBQUMsQ0FBQyxDQUNkLENBQUM7UUFDSCx5QkFBeUI7UUFDekIsT0FBTyxTQUFTLENBQUM7SUFDbkIsQ0FBQztJQXZCZSwwQkFBWSxlQXVCM0IsQ0FBQTtJQUVELFNBQWdCLGVBQWU7UUFDN0IsT0FBTyxhQUFhLENBQUMsWUFBWSxFQUFFLEdBQUcsaUJBQWlCLENBQUM7SUFDMUQsQ0FBQztJQUZlLDZCQUFlLGtCQUU5QixDQUFBO0lBRUQsU0FB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ZCBkaXNhc3NvY2lhdGluZyBzaW5jZSB0aGlzIGlzIG5vdCBzdXBwb3J0ZWQgYnkgdGhlIFhybS5XZWJBcGkgY2xpZW50IHNpZGUgYXBpIGF0IHRoaXMgdGltZVxyXG5uYW1lc3BhY2UgV2ViQXBpUmVxdWVzdCB7XHJcbiAgdmFyIHdlYkFwaVVybDogc3RyaW5nID0gXCJcIjtcclxuICBleHBvcnQgZnVuY3Rpb24gZ2V0V2ViQXBpVXJsKCkge1xyXG4gICAgdmFyIGNvbnRleHQ6IFhybS5HbG9iYWxDb250ZXh0O1xyXG4gICAgdmFyIGNsaWVudFVybDogc3RyaW5nO1xyXG4gICAgdmFyIGFwaVZlcnNpb246IHN0cmluZztcclxuICAgIGlmICh3ZWJBcGlVcmwpIHJldHVybiB3ZWJBcGlVcmw7XHJcblxyXG4gICAgaWYgKEdldEdsb2JhbENvbnRleHQpIHtcclxuICAgICAgY29udGV4dCA9IEdldEdsb2JhbENvbnRleHQoKTtcclxuICAgIH0gZWxzZSB7XHJcbiAgICAgIGlmIChYcm0pIHtcclxuICAgICAgICBjb250ZXh0ID0gWHJtLlBhZ2UuY29udGV4dDtcclxuICAgICAgfSBlbHNlIHtcclxuICAgICAgICB0aHJvdyBuZXcgRXJyb3IoXCJDb250ZXh0IGlzIG5vdCBhdmFpbGFibGUuXCIpO1xyXG4gICAgICB9XHJcbiAgICB9XHJcbiAgICBjbGllbnRVcmwgPSBjb250ZXh0LmdldENsaWVudFVybCgpO1xyXG4gICAgdmFyIHZlcnNpb25QYXJ0cyA9IGNvbnRleHQuZ2V0VmVyc2lvbigpLnNwbGl0KFwiLlwiKTtcclxuXHJcbiAgICAgd2ViQXBpVXJsID0gYCR7Y2xpZW50VXJsfS9hcGkvZGF0YS92JHt2ZXJzaW9uUGFydHNbMF19LiR7XHJcbiAgICAgIHZlcnNpb25QYXJ0c1sxXVxyXG4gICAgfWA7XHJcbiAgICAvLyBBZGQgdGhlIFdlYkFwaSB2ZXJzaW9uXHJcbiAgICByZXR1cm4gd2ViQXBpVXJsO1xyXG4gIH1cclxuICBcclxuICBleHBvcnQgZnVuY3Rpb24gZ2V0T2RhdGFDb250ZXh0KCkge1xyXG4gICAgcmV0dXJuIFdlYkFwaVJlcXVlc3QuZ2V0V2ViQXBpVXJsKCkgKyBcIi8kbWV0YWRhdGEjJHJlZlwiO1xyXG4gIH1cclxuXHJcbiAgZXhwb3J0IGZ1bmN0aW9uIHJlcXVlc3QoXHJcbiAgICBhY3Rpb246IFwiUE9TVFwiIHwgXCJQQVRDSFwiIHwgXCJQVVRcIiB8IFwiR0VUXCIgfCBcIkRFTEVURVwiLFxyXG4gICAgdXJpOiBzdHJpbmcsXHJcbiAgICBwYXlsb2FkPzogYW55LFxyXG4gICAgaW5jbHVkZUZvcm1hdHRlZFZhbHVlcz86IGJvb2xlYW4sXHJcbiAgICBtYXhQYWdlU2l6ZT86IG51bWJlclxyXG4gICkge1xyXG4gICAgLy8gQ29uc3RydWN0IGEgZnVsbHkgcXVhbGlmaWVkIFVSSSBpZiBhIHJlbGF0aXZlIFVSSSBpcyBwYXNzZWQgaW4uXHJcbiAgICBpZiAodXJpLmNoYXJBdCgwKSA9PT0gXCIvXCIpIHtcclxuICAgICAgdXJpID0gV2ViQXBpUmVxdWVzdC5nZXRXZWJBcGlVcmwoKSArIHVyaTtcclxuICAgIH1cclxuXHJcbiAgICByZXR1cm4gbmV3IFByb21pc2UoZnVuY3Rpb24ocmVzb2x2ZSwgcmVqZWN0KSB7XHJcbiAgICAgIHZhciByZXF1ZXN0ID0gbmV3IFhNTEh0dHBSZXF1ZXN0KCk7XHJcbiAgICAgIHJlcXVlc3Qub3BlbihhY3Rpb24sIGVuY29kZVVSSSh1cmkpLCB0cnVlKTtcclxuICAgICAgcmVxdWVzdC5zZXRSZXF1ZXN0SGVhZGVyKFwiT0RhdGEtTWF4VmVyc2lvblwiLCBcIjQuMFwiKTtcclxuICAgICAgcmVxdWVzdC5zZXRSZXF1ZXN0SGVhZGVyKFwiT0RhdGEtVmVyc2lvblwiLCBcIjQuMFwiKTtcclxuICAgICAgcmVxdWVzdC5zZXRSZXF1ZXN0SGVhZGVyKFwiQWNjZXB0XCIsIFwiYXBwbGljYXRpb24vanNvblwiKTtcclxuICAgICAgcmVxdWVzdC5zZXRSZXF1ZXN0SGVhZGVyKFxyXG4gICAgICAgIFwiQ29udGVudC1UeXBlXCIsXHJcbiAgICAgICAgXCJhcHBsaWNhdGlvbi9qc29uOyBjaGFyc2V0PXV0Zi04XCJcclxuICAgICAgKTtcclxuICAgICAgaWYgKG1heFBhZ2VTaXplKSB7XHJcbiAgICAgICAgcmVxdWVzdC5zZXRSZXF1ZXN0SGVhZGVyKFwiUHJlZmVyXCIsIFwib2RhdGEubWF4cGFnZXNpemU9XCIgKyBtYXhQYWdlU2l6ZSk7XHJcbiAgICAgIH1cclxuICAgICAgaWYgKGluY2x1ZGVGb3JtYXR0ZWRWYWx1ZXMpIHtcclxuICAgICAgICByZXF1ZXN0LnNldFJlcXVlc3RIZWFkZXIoXHJcbiAgICAgICAgICBcIlByZWZlclwiLFxyXG4gICAgICAgICAgXCJvZGF0YS5pbmNsdWRlLWFubm90YXRpb25zPU9EYXRhLkNvbW11bml0eS5EaXNwbGF5LlYxLkZvcm1hdHRlZFZhbHVlXCJcclxuICAgICAgICApO1xyXG4gICAgICB9XHJcbiAgICAgIHJlcXVlc3Qub25yZWFkeXN0YXRlY2hhbmdlID0gZnVuY3Rpb24oKSB7XHJcbiAgICAgICAgaWYgKHRoaXMucmVhZHlTdGF0ZSA9PT0gNCkge1xyXG4gICAgICAgICAgcmVxdWVzdC5vbnJlYWR5c3RhdGVjaGFuZ2UgPSBudWxsO1xyXG4gICAgICAgICAgc3dpdGNoICh0aGlzLnN0YXR1cykge1xyXG4gICAgICAgICAgICBjYXNlIDIwMDogLy8gU3VjY2VzcyB3aXRoIGNvbnRlbnQgcmV0dXJuZWQgaW4gcmVzcG9uc2UgYm9keS5cclxuICAgICAgICAgICAgY2FzZSAyMDQ6IC8vIFN1Y2Nlc3Mgd2l0aCBubyBjb250ZW50IHJldHVybmVkIGluIHJlc3BvbnNlIGJvZHkuXHJcbiAgICAgICAgICAgICAgcmVzb2x2ZSh0aGlzKTtcclxuICAgICAgICAgICAgICBicmVhaztcclxuICAgICAgICAgICAgZGVmYXVsdDpcclxuICAgICAgICAgICAgICAvLyBBbGwgb3RoZXIgc3RhdHVzZXMgYXJlIHVuZXhwZWN0ZWQgc28gYXJlIHRyZWF0ZWQgbGlrZSBlcnJvcnMuXHJcbiAgICAgICAgICAgICAgdmFyIGVycm9yO1xyXG4gICAgICAgICAgICAgIHRyeSB7XHJcbiAgICAgICAgICAgICAgICBlcnJvciA9IEpTT04ucGFyc2UocmVxdWVzdC5yZXNwb25zZSkuZXJyb3I7XHJcbiAgICAgICAgICAgICAgfSBjYXRjaCAoZSkge1xyXG4gICAgICAgICAgICAgICAgZXJyb3IgPSBuZXcgRXJyb3IoXCJVbmV4cGVjdGVkIEVycm9yXCIpO1xyXG4gICAgICAgICAgICAgIH1cclxuICAgICAgICAgICAgICByZWplY3QoZXJyb3IpO1xyXG4gICAgICAgICAgICAgIGJyZWFrO1xyXG4gICAgICAgICAgfVxyXG4gICAgICAgIH1cclxuICAgICAgfTtcclxuICAgICAgcmVxdWVzdC5zZW5kKEpTT04uc3RyaW5naWZ5KHBheWxvYWQpKTtcclxuICAgIH0pO1xyXG4gIH1cclxufVxyXG4iLCIvLyBUaGlzIGlzIHJlcXVpcmVkIGR1ZSB0byBhIGJ1ZyB3aGVyZSBlbnRpdHkgbWV0YWRhdGEgaXMgbm90IGxvYWRlZCBjb3JyZWN0bHkgaW4gd2VicmVzb3VyY2VzIHRoYXQgb3BlbiBpbiBhIG5ldyB3aW5kb3dcclxuLy8gTm90IG5lZWRlZCBpZiB0aGUgd2VicmVzb3VyY2UgaXMgZW1iZGVkZGVkIGluc2lkZSBhIGZvcm1cclxudmFyIHdpbmRvd1N0YXRpYyA6YW55PSB3aW5kb3c7XHJcbndpbmRvd1N0YXRpYy5FTlRJVFlfU0VUX05BTUVTID0gYHtcclxuICAgIFwiYWNjb3VudFwiOiBcImFjY291bnRzXCIsXHJcbiAgICBcImNvbnRhY3RcIjogXCJjb250YWN0c1wiLFxyXG4gICAgXCJvcHBvcnR1bml0eVwiOiBcIm9wcG9ydHVuaXRpZXNcIlxyXG59YDtcclxuXHJcbndpbmRvd1N0YXRpYy5FTlRJVFlfUFJJTUFSWV9LRVlTID0gYHtcclxuICAgIFwiYWNjb3VudFwiIDogXCJhY2NvdW50aWRcIixcclxuICAgIFwiY29udGFjdFwiIDogXCJjb250YWN0aWRcIixcclxuICAgIFwib3Bwb3J0dW5pdHlcIiA6IFwib3Bwb3J0dW5pdHlpZFwiXHJcbn1gO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWVzOlxyXG4vLyAgVXNpbmcgdGhlIEJvdW5kIGZ1bmN0aW9uIEFkZFRvUXVldWVSZXNwb25zZVxyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvdXNlLXdlYi1hcGktYWN0aW9ucyNib3VuZC1hY3Rpb25zXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuXHJcbiAgaXQoXCJBZGRUb1F1ZXVlUmVzcG9uc2VcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG4gICAgLy8gQ3JlYXRlIFF1ZXVlXHJcbiAgICB2YXIgcXVldWVpZCA9ICg8YW55Pihhd2FpdCBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcInF1ZXVlXCIse1wibmFtZVwiIDogXCJTYW1wbGUgUXVldWVcIn0pKSkuaWQ7XHJcbiAgICBcclxuICAgIC8vIENyZWF0ZSBsZXR0ZXJcclxuICAgIHZhciBsZXR0ZXJpZCA9ICg8YW55Pihhd2FpdCBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImxldHRlclwiLHtcInN1YmplY3RcIiA6IFwiU2FtcGxlIExldHRlclwifSkpKS5pZDtcclxuXHJcbiAgICB0cnl7XHJcbiAgICAvLyBFeGVjdXRlIHJlcXVlc3RcclxuICAgIHZhciBBZGRUb1F1ZXVlUmVxdWVzdCA9IG5ldyBjbGFzcyB7XHJcbiAgICAgIGVudGl0eSA9IHtcclxuICAgICAgICBpZDogcXVldWVpZCxcclxuICAgICAgICBlbnRpdHlUeXBlOiBcInF1ZXVlXCJcclxuICAgICAgfTtcclxuICAgICAgVGFyZ2V0ID0ge1xyXG4gICAgICAgICAgaWQ6IGxldHRlcmlkLFxyXG4gICAgICAgICAgZW50aXR5VHlwZTogXCJsZXR0ZXJcIlxyXG4gICAgICB9O1xyXG5cclxuICAgICAgZ2V0TWV0YWRhdGEoKTogYW55IHtcclxuICAgICAgICByZXR1cm4ge1xyXG5cdFx0Ym91bmRQYXJhbWV0ZXI6IFwiZW50aXR5XCIsXHJcblx0XHRwYXJhbWV0ZXJUeXBlczoge1xyXG5cdFx0XHRcImVudGl0eVwiOiB7XHJcblx0XHRcdFx0dHlwZU5hbWU6IFwibXNjcm0ucXVldWVcIixcclxuICAgICAgICAgICAgICAgIHN0cnVjdHVyYWxQcm9wZXJ0eTogNVxyXG5cdFx0XHR9LFx0XHRcclxuXHRcdFx0XCJRdWV1ZUl0ZW1Qcm9wZXJ0aWVzXCI6IHtcclxuXHRcdFx0XHR0eXBlTmFtZTogXCJtc2NybS5xdWV1ZWl0ZW1cIixcclxuICAgICAgICBzdHJ1Y3R1cmFsUHJvcGVydHk6IDVcclxuXHRcdFx0fSxcdFx0XHJcblx0XHRcdFwiU291cmNlUXVldWVcIjoge1xyXG5cdFx0XHRcdHR5cGVOYW1lOiBcIm1zY3JtLnF1ZXVlXCIsXHJcbiAgICAgICAgc3RydWN0dXJhbFByb3BlcnR5OiA1XHJcblx0XHRcdH0sXHRcdFxyXG5cdFx0XHRcIlRhcmdldFwiOiB7XHJcblx0XHRcdFx0dHlwZU5hbWU6IFwibXNjcm0uY3JtYmFzZWVudGl0eVwiLFxyXG4gICAgICAgICAgICAgICAgc3RydWN0dXJhbFByb3BlcnR5OiA1XHJcblx0XHRcdH0sXHRcdFxyXG5cdFx0fSxcclxuXHRcdG9wZXJhdGlvblR5cGU6IDAsXHJcblx0XHRvcGVyYXRpb25OYW1lOiBcIkFkZFRvUXVldWVcIlxyXG5cdH07XHJcbiAgICAgIH1cclxuICAgIH0oKTtcclxuXHJcblxyXG4gICAgdmFyIHJlc3BvbnNlIDoge1xyXG4gICAgICBRdWV1ZUl0ZW1JZDogc3RyaW5nXHJcbiAgICB9PSBhd2FpdCAoPGFueT4oYXdhaXQgWHJtLldlYkFwaS5vbmxpbmUuZXhlY3V0ZShBZGRUb1F1ZXVlUmVxdWVzdCkpKS5qc29uKCk7XHJcbiAgXHJcbiAgICBhc3NlcnQuaXNTdHJpbmcocmVzcG9uc2UuUXVldWVJdGVtSWQsXCJRdWV1ZUl0ZW1JZCByZXR1cm5lZFwiKTtcclxuXHJcbiAgICB9XHJcbiAgICBmaW5hbGx5e1xyXG4gICAgICAvLyBEZWxldGUgTGV0dGVyXHJcbiAgICAgIGlmIChsZXR0ZXJpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwibGV0dGVyXCIsIGxldHRlcmlkKTtcclxuICAgICAgfVxyXG4gICAgICAgLy8gRGVsZXRlIFF1ZXVlXHJcbiAgICAgICBpZiAocXVldWVpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwicXVldWVcIiwgcXVldWVpZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICAgIFxyXG5cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZXM6XHJcbi8vICAxLiBDcmVhdGluZyBhbiBvcHBvcnR1bml0eVxyXG4vLyAgMi4gV2lubmluZyBhbiBvcHBvcnR1bml0eSB1c2luZyB0aGUgZXhlY3V0ZSBtZXRob2RcclxuLy8gU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvdXNlLXdlYi1hcGktZnVuY3Rpb25zXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuXHJcbiAgaXQoXCJXaW4gT3Bwb3J0dW5pdHlcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG4gICAgXHJcbiAgICB2YXIgYWNjb3VudDE6IGFueSA9IHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiXHJcbiAgICB9O1xyXG5cclxuICAgIHZhciBjcmVhdGVBY2NvdW50UmVzcG9uc2U6IHtcclxuICAgICAgZW50aXR5VHlwZTogU3RyaW5nO1xyXG4gICAgICBpZDogU3RyaW5nO1xyXG4gICAgfSA9IDxhbnk+YXdhaXQgWHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIGFjY291bnQxKTtcclxuXHJcbiAgICBhY2NvdW50MS5hY2NvdW50aWQgPSBjcmVhdGVBY2NvdW50UmVzcG9uc2UuaWQ7XHJcblxyXG4gICAgdmFyIG9wcG9ydHVuaXR5MTogYW55ID0ge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBPcHBvcnR1bml0eVwiLFxyXG4gICAgICBlc3RpbWF0ZWR2YWx1ZTogMTAwMCxcclxuICAgICAgZXN0aW1hdGVkY2xvc2VkYXRlOiBcIjIwMTktMDItMTBcIixcclxuICAgICAgXCJwYXJlbnRhY2NvdW50aWRAb2RhdGEuYmluZFwiOiBgYWNjb3VudHMoJHthY2NvdW50MS5hY2NvdW50aWR9KWBcclxuICAgIH07XHJcblxyXG4gICAgdmFyIGNyZWF0ZU9wcG9ydHVuaXR5UmVzcG9uc2U6IHtcclxuICAgICAgZW50aXR5VHlwZTogU3RyaW5nO1xyXG4gICAgICBpZDogU3RyaW5nO1xyXG4gICAgfSA9IDxhbnk+YXdhaXQgWHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJvcHBvcnR1bml0eVwiLCBvcHBvcnR1bml0eTEpO1xyXG5cclxuICAgIG9wcG9ydHVuaXR5MS5vcHBvcnR1bml0eWlkID0gY3JlYXRlT3Bwb3J0dW5pdHlSZXNwb25zZS5pZDtcclxuXHJcbiAgICAvLyBFeGVjdXRlIHJlcXVlc3RcclxuICAgIHZhciB3aW5PcHBvcnR1bml0eVJlcXVlc3QgPSBuZXcgY2xhc3Mge1xyXG4gICAgICBPcHBvcnR1bml0eUNsb3NlID0ge1xyXG4gICAgICAgIGRlc2NyaXB0aW9uOiBcIlNhbXBsZSBPcHBvcnR1bml0eSBDbG9zZVwiLFxyXG4gICAgICAgIHN1YmplY3Q6IFwiU2FtcGxlXCIsXHJcbiAgICAgICAgXCJAb2RhdGEudHlwZVwiOiBcIk1pY3Jvc29mdC5EeW5hbWljcy5DUk0ub3Bwb3J0dW5pdHljbG9zZVwiLFxyXG4gICAgICAgIFwib3Bwb3J0dW5pdHlpZEBvZGF0YS5iaW5kXCI6IGBvcHBvcnR1bml0aWVzKCR7XHJcbiAgICAgICAgICBvcHBvcnR1bml0eTEub3Bwb3J0dW5pdHlpZFxyXG4gICAgICAgIH0pYFxyXG4gICAgICB9O1xyXG4gICAgICBTdGF0dXMgPSAzO1xyXG5cclxuICAgICAgZ2V0TWV0YWRhdGEoKTogYW55IHtcclxuICAgICAgICByZXR1cm4ge1xyXG4gICAgICAgICAgcGFyYW1ldGVyVHlwZXM6IHtcclxuICAgICAgICAgICAgT3Bwb3J0dW5pdHlDbG9zZToge1xyXG4gICAgICAgICAgICAgIHR5cGVOYW1lOiBcIm1zY3JtLm9wcG9ydHVuaXR5Y2xvc2VcIixcclxuICAgICAgICAgICAgICBzdHJ1Y3R1cmFsUHJvcGVydHk6IDVcclxuICAgICAgICAgICAgfSxcclxuICAgICAgICAgICAgU3RhdHVzOiB7XHJcbiAgICAgICAgICAgICAgdHlwZU5hbWU6IFwiRWRtLkludDMyXCIsXHJcbiAgICAgICAgICAgICAgc3RydWN0dXJhbFByb3BlcnR5OiAxXHJcbiAgICAgICAgICAgIH1cclxuICAgICAgICAgIH0sXHJcbiAgICAgICAgICBvcGVyYXRpb25UeXBlOiAwLFxyXG4gICAgICAgICAgb3BlcmF0aW9uTmFtZTogXCJXaW5PcHBvcnR1bml0eVwiXHJcbiAgICAgICAgfTtcclxuICAgICAgfVxyXG4gICAgfSgpO1xyXG4gICAgdmFyIHJhd1Jlc3BvbnNlID0gPGFueT4oXHJcbiAgICAgIGF3YWl0IFhybS5XZWJBcGkub25saW5lLmV4ZWN1dGUod2luT3Bwb3J0dW5pdHlSZXF1ZXN0KVxyXG4gICAgKTtcclxuICAgIHZhciByZXNwb25zZSA9IGF3YWl0IHJhd1Jlc3BvbnNlLnRleHQoKTtcclxuXHJcbiAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgaWYgKHJlc3BvbnNlLmlkKSB7XHJcbiAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiYWNjb3VudFwiLCByZXNwb25zZS5pZCk7XHJcbiAgICB9XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWU6XHJcbi8vICBDcmVhdGluZyBhIHJlY29yZFxyXG5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJDcmVhdGVcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG4gICAgdmFyIGNvbnRleHQgPSBHZXRHbG9iYWxDb250ZXh0KCk7XHJcblxyXG4gICAgdmFyIGFjY291bnQxOiB7XHJcbiAgICAgIGFjY291bnRpZD86IHN0cmluZztcclxuICAgICAgbmFtZT86IHN0cmluZztcclxuICAgICAgYWNjb3VudGNhdGVnb3J5Y29kZTogTnVtYmVyOyAvL09wdGlvbnNldFxyXG4gICAgICBjcmVkaXRsaW1pdDogTnVtYmVyOyAvLyBNb25leVxyXG4gICAgICBjcmVkaXRvbmhvbGQ6IEJvb2xlYW47IC8vIEJvb2xlYW5cclxuICAgICAgbnVtYmVyb2ZlbXBsb3llZXM6IE51bWJlcjsgLy8gSW50ZWdlclxyXG4gICAgICBsYXN0b25ob2xkdGltZTogRGF0ZTsgLy8gRGF0ZVxyXG4gICAgICBbaW5kZXg6IHN0cmluZ106IGFueTsgLy8gQWxsb3cgc2V0dGluZyBAb2RhdGEgcHJvcGVydGllc1xyXG4gICAgfTtcclxuXHJcbiAgICBhY2NvdW50MSA9IHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiLFxyXG4gICAgICBhY2NvdW50Y2F0ZWdvcnljb2RlOiAxLCAvL1ByZWZlcnJlZF9DdXN0b21lclxyXG4gICAgICBjcmVkaXRsaW1pdDogMTAwMCxcclxuICAgICAgY3JlZGl0b25ob2xkOiB0cnVlLFxyXG4gICAgICBudW1iZXJvZmVtcGxveWVlczogMTAsXHJcbiAgICAgIGxhc3RvbmhvbGR0aW1lOiBuZXcgRGF0ZSgpLFxyXG4gICAgICBcInByZWZlcnJlZHN5c3RlbXVzZXJpZEBvZGF0YS5iaW5kXCI6IGBzeXN0ZW11c2Vycygke2NvbnRleHRcclxuICAgICAgICAuZ2V0VXNlcklkKClcclxuICAgICAgICAucmVwbGFjZShcIntcIiwgXCJcIilcclxuICAgICAgICAucmVwbGFjZShcIn1cIiwgXCJcIil9KWBcclxuICAgIH07XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgLy8gQ3JlYXRlIEFjY291bnRcclxuICAgICAgYWNjb3VudDEuYWNjb3VudGlkID0gKGF3YWl0ICg8YW55PihcclxuICAgICAgICBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudDEpXHJcbiAgICAgICkpKS5pZDtcclxuXHJcbiAgICAgIGlmICghYWNjb3VudDEuYWNjb3VudGlkKSB7XHJcbiAgICAgICAgdGhyb3cgbmV3IEVycm9yKFwiQWNjb3VudCBub3QgY3JlYXRlZFwiKTtcclxuICAgICAgfVxyXG5cclxuICAgICAgLy8gQ2hlY2sgdGhlIGFjY291bnQgaGFzIGJlZW4gY3JlYXRlZFxyXG4gICAgICB2YXIgYWNjb3VudDIgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFxyXG4gICAgICAgIFwiYWNjb3VudFwiLFxyXG4gICAgICAgIGFjY291bnQxLmFjY291bnRpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PW5hbWVcIlxyXG4gICAgICApO1xyXG5cclxuICAgICAgYXNzZXJ0LmVxdWFsKGFjY291bnQyLm5hbWUsIFwiU2FtcGxlIEFjY291bnRcIiwgXCJBY2NvdW50IGNyZWF0ZWRcIik7XHJcbiAgICB9IGZpbmFsbHkge1xyXG4gICAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgICBpZiAoYWNjb3VudDEuYWNjb3VudGlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJhY2NvdW50XCIsIGFjY291bnQxLmFjY291bnRpZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZTpcclxuLy8gIENyZWF0aW5nIGFuIGFjY291bnQgYW5kIDIgcmVsYXRlZCBjb250YWN0cyBpbiB0aGUgc2FtZSB0cmFuc2FjdGlvblxyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvYXNzb2NpYXRlLWRpc2Fzc29jaWF0ZS1lbnRpdGllcy11c2luZy13ZWItYXBpI2Fzc29jaWF0ZS1lbnRpdGllcy1vbi1jcmVhdGVcclxuXHJcbmRlc2NyaWJlKFwiXCIsIGZ1bmN0aW9uKCkge1xyXG4gIGl0KFwiRGVlcCBJbnNlcnRcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG4gICAgdmFyIGNvbnRleHQgPSBHZXRHbG9iYWxDb250ZXh0KCk7XHJcbiAgICB2YXIgYWNjb3VudDoge1xyXG4gICAgICBhY2NvdW50aWQ/OiBzdHJpbmc7XHJcbiAgICAgIG5hbWU/OiBzdHJpbmc7XHJcbiAgICAgIGNvbnRhY3RfY3VzdG9tZXJfYWNjb3VudHMgOiB7XHJcbiAgICAgICAgZmlyc3RuYW1lOiBzdHJpbmc7XHJcbiAgICAgICAgbGFzdG5hbWU6IHN0cmluZztcclxuICAgICAgfVtdO1xyXG4gICAgfTtcclxuXHJcbiAgICBhY2NvdW50ID0ge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBBY2NvdW50XCIsXHJcbiAgICAgIGNvbnRhY3RfY3VzdG9tZXJfYWNjb3VudHM6IFtcclxuICAgICAgICB7XHJcbiAgICAgICAgICBmaXJzdG5hbWU6IFwiU2FtcGxlXCIsXHJcbiAgICAgICAgICBsYXN0bmFtZTogXCJjb250YWN0IDFcIlxyXG4gICAgICAgIH0sXHJcbiAgICAgICAge1xyXG4gICAgICAgICAgZmlyc3RuYW1lOiBcIlNhbXBsZVwiLFxyXG4gICAgICAgICAgbGFzdG5hbWU6IFwiQ29udGFjdCAyXCJcclxuICAgICAgICB9LFxyXG4gICAgICBdXHJcbiAgICB9O1xyXG5cclxuICAgIHRyeSB7XHJcbiAgICAgIC8vIENyZWF0ZSBBY2NvdW50ICYgQ29udGFjdHNcclxuICAgICAgYWNjb3VudC5hY2NvdW50aWQgPSAoYXdhaXQgKDxhbnk+KFxyXG4gICAgICAgIFhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiYWNjb3VudFwiLCBhY2NvdW50KVxyXG4gICAgICApKSkuaWQ7XHJcblxyXG4gICAgICBpZiAoIWFjY291bnQuYWNjb3VudGlkKVxyXG4gICAgICAgIHRocm93IG5ldyBFcnJvcihcIkFjY291bnQgbm90IGNyZWF0ZWRcIik7XHJcblxyXG4gICAgICB2YXIgYWNjb3VudENyZWF0ZWQgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFwiYWNjb3VudFwiLGFjY291bnQuYWNjb3VudGlkLFwiPyRzZWxlY3Q9bmFtZSYkZXhwYW5kPWNvbnRhY3RfY3VzdG9tZXJfYWNjb3VudHMoJHNlbGVjdD1maXJzdG5hbWUsbGFzdG5hbWUpXCIpO1xyXG4gICAgIFxyXG4gICAgICBhc3NlcnQuZXF1YWwoYWNjb3VudENyZWF0ZWQuY29udGFjdF9jdXN0b21lcl9hY2NvdW50cy5sZW5ndGgsIDIsIFwiQWNjb3VudCBjcmVhdGVkIHdpdGggMiBjb250YWN0c1wiKTtcclxuXHJcbiAgICB9IGZpbmFsbHkge1xyXG4gICAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgICBpZiAoYWNjb3VudC5hY2NvdW50aWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudC5hY2NvdW50aWQpO1xyXG4gICAgICB9XHJcbiAgICB9XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWU6XHJcbi8vICBEZWxldGluZyBhIHJlY29yZFxyXG5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJEZWxldGVcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG5cclxuICAgIHZhciByZWNvcmQ6IHtcclxuICAgICAgYWNjb3VudGlkPzogc3RyaW5nO1xyXG4gICAgICBuYW1lPzogc3RyaW5nO1xyXG4gICAgfTtcclxuICAgIFxyXG4gICAgcmVjb3JkID0gIHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiXHJcbiAgICB9O1xyXG5cclxuICAgIC8vIENyZWF0ZSBBY2NvdW50XHJcbiAgICByZWNvcmQuYWNjb3VudGlkID0gKGF3YWl0ICg8YW55PihcclxuICAgICAgWHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIHJlY29yZClcclxuICAgICkpKS5pZDtcclxuXHJcbiAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgaWYgKHJlY29yZC5hY2NvdW50aWQpIHtcclxuICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJhY2NvdW50XCIsIHJlY29yZC5hY2NvdW50aWQpO1xyXG4gICAgfVxyXG5cclxuICAgIC8vIENoZWNrIHRoZSBhY2NvdW50IGhhcyBiZWVuIGRlbGV0ZWRcclxuICAgIHZhciBmZXRjaCA9IGA8ZmV0Y2ggbm8tbG9jaz1cInRydWVcIiA+XHJcbiAgICAgICA8ZW50aXR5IG5hbWU9XCJhY2NvdW50XCIgPlxyXG4gICAgICAgICA8ZmlsdGVyPlxyXG4gICAgICAgICAgIDxjb25kaXRpb24gYXR0cmlidXRlPVwiYWNjb3VudGlkXCIgb3BlcmF0b3I9XCJlcVwiIHZhbHVlPVwiJHtcclxuICAgICAgICAgICAgIHJlY29yZC5hY2NvdW50aWRcclxuICAgICAgICAgICB9XCIgLz5cclxuICAgICAgICAgPC9maWx0ZXI+XHJcbiAgICAgICA8L2VudGl0eT5cclxuICAgICA8L2ZldGNoPmA7XHJcblxyXG4gICAgdmFyIGFjY291bnRzID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZU11bHRpcGxlUmVjb3JkcyhcclxuICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgIFwiP2ZldGNoWG1sPVwiICsgZmV0Y2hcclxuICAgICk7XHJcblxyXG4gICAgYXNzZXJ0LmVxdWFsKGFjY291bnRzLmVudGl0aWVzLmxlbmd0aCwgMCwgXCJBY2NvdW50IGRlbGV0ZWRcIik7XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWU6XHJcbi8vICBDcmVhdGluZyBhIHJlY29yZCBhbmQgdGhlbiByZXRyaWV2aW5nIGl0IHVuY2hhbmdlZCB3aXRoIHRoZSBzYW1lIGV0YWdcclxuLy8gIFRoaXMgaXMgYSB0ZWNobmlxdWUgZW1wbG95ZWQgYnkgdGhlIGNsaWVudCBzaWRlIGFwaSB0byBhdm9pZCB1bmVjY2VzYXJ5IGRhdGEgYmVpbmcgdHJhbnNmZXJyZWQgb3ZlciB0aGUgbmV0d29ya1xyXG4vLyAgVGhpcyBzYW1wbGUgc2hvd3MgaG93IHZhcnlpbmcgdGhlIG9wdGlvbnMgd2lsbCBpbnZhbGlkYXRlIHRoZSBjbGllbnQgc2lkZSBjYWNoZVxyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvY3JlYXRlLWVudGl0eS13ZWItYXBpXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcImV0YWdzIHdpdGggJGV4cGFuZFwiLCBhc3luYyBmdW5jdGlvbigpIHtcclxuICAgIHRoaXMudGltZW91dCg5MDAwMCk7XHJcblxyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG5cclxuICAgIC8vIENyZWF0ZSBhIGNvbnRhY3RcclxuICAgIHZhciBjb250YWN0aWQ6IHtcclxuICAgICAgZW50aXR5VHlwZTogc3RyaW5nO1xyXG4gICAgICBpZDogc3RyaW5nO1xyXG4gICAgfSA9IGF3YWl0ICg8YW55PlhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiY29udGFjdFwiLCB7XHJcbiAgICAgIGxhc3RuYW1lOiBcIlNhbXBsZSBDb250YWN0XCJcclxuICAgIH0pKTtcclxuXHJcbiAgICAvLyBDcmVhdGUgYWNjb3VudFxyXG4gICAgdmFyIGFjY291bnRpZDoge1xyXG4gICAgICBlbnRpdHlUeXBlOiBzdHJpbmc7XHJcbiAgICAgIGlkOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgKDxhbnk+WHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiLFxyXG4gICAgICBcInByaW1hcnljb250YWN0aWRAb2RhdGEuYmluZFwiOiBgL2NvbnRhY3RzKCR7Y29udGFjdGlkLmlkfSlgXHJcbiAgICB9KSk7XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgLy8gUmVhZCB0aGUgQWNjb3VudFxyXG4gICAgICB2YXIgcXVlcnkxID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZVJlY29yZChcclxuICAgICAgICBcImFjY291bnRcIixcclxuICAgICAgICBhY2NvdW50aWQuaWQsXHJcbiAgICAgICAgXCI/JHNlbGVjdD1uYW1lJiRleHBhbmQ9cHJpbWFyeWNvbnRhY3RpZFwiXHJcbiAgICAgICk7XHJcbiAgICAgIHZhciBldGFnMSA9IHF1ZXJ5MVtcIkBvZGF0YS5ldGFnXCJdO1xyXG5cclxuICAgICAgLy8gUmVhZCB0aGUgQWNjb3VudCBhZ2FpbiAoVGhpcyB3aWxsIHJldHVybiB0aGUgc2FtZSByZWNvcmQgYXMgYWJvdmUgc2luY2UgdGhlIHNlcnZlciB3aWxsIHJldHVybiAzMDQgTm90IE1vZGlmaWVkIClcclxuICAgICAgdmFyIHF1ZXJ5MiA9IGF3YWl0IFhybS5XZWJBcGkucmV0cmlldmVSZWNvcmQoXHJcbiAgICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgICAgYWNjb3VudGlkLmlkLFxyXG4gICAgICAgIFwiPyRzZWxlY3Q9bmFtZSYkZXhwYW5kPXByaW1hcnljb250YWN0aWRcIlxyXG4gICAgICApO1xyXG4gICAgICB2YXIgZXRhZzIgPSBxdWVyeTJbXCJAb2RhdGEuZXRhZ1wiXTtcclxuXHJcbiAgICAgIGFzc2VydC5lcXVhbChldGFnMSwgZXRhZzIsIFwiUmVjb3JkIG5vdCBtb2RpZmllZFwiKTtcclxuICAgICAgYXNzZXJ0LmVxdWFsKFxyXG4gICAgICAgIFwiU2FtcGxlIENvbnRhY3RcIixcclxuICAgICAgICBxdWVyeTIucHJpbWFyeWNvbnRhY3RpZC5sYXN0bmFtZSxcclxuICAgICAgICBcIlJlbGF0ZWQgY29udGFjdCByZXR1cm5lZFwiXHJcbiAgICAgICk7XHJcblxyXG4gICAgICAvLyBVcGRhdGUgdGhlIGNvbnRhY3QgbmFtZVxyXG4gICAgICBhd2FpdCAoPGFueT5Ycm0uV2ViQXBpLnVwZGF0ZVJlY29yZChcImNvbnRhY3RcIiwgY29udGFjdGlkLmlkLCB7XHJcbiAgICAgICAgbGFzdG5hbWU6IFwiU2FtcGxlIENvbnRhY3QgKGVkaXRlZClcIlxyXG4gICAgICB9KSk7XHJcblxyXG4gICAgICAvLyBSZWFkIHRoZSBBY2NvdW50IGFnYWluLiBTaW5jZSBvbmx5IHRoZSByZWxhdGVkIGV4cGFuZGVkIHJlY29yZCB3YXMgdXBkYXRlZCwgdGhlIHJldHJpZXZlIHdpbGwgbm90IHJldHVybiB0aGUgY29ycmVjdCB2YWx1ZS5cclxuICAgICAgdmFyIHF1ZXJ5MyA9IGF3YWl0IFhybS5XZWJBcGkucmV0cmlldmVSZWNvcmQoXHJcbiAgICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgICAgYWNjb3VudGlkLmlkLFxyXG4gICAgICAgIFwiPyRzZWxlY3Q9bmFtZSYkZXhwYW5kPXByaW1hcnljb250YWN0aWRcIlxyXG4gICAgICApO1xyXG4gICAgICB2YXIgZXRhZzMgPSBxdWVyeTNbXCJAb2RhdGEuZXRhZ1wiXTtcclxuXHJcbiAgICAgIGFzc2VydC5lcXVhbChldGFnMSwgZXRhZzMsIFwiUmVjb3JkIG5vdCBtb2RpZmllZFwiKTtcclxuICAgICAgYXNzZXJ0LmVxdWFsKFxyXG4gICAgICAgIFwiU2FtcGxlIENvbnRhY3RcIixcclxuICAgICAgICBxdWVyeTMucHJpbWFyeWNvbnRhY3RpZC5sYXN0bmFtZSxcclxuICAgICAgICBcIlVuY2hhbmdlZCBjb250YWN0XCJcclxuICAgICAgKTtcclxuXHJcbiAgICAgIC8vIFdvcmthcm91bmQ6IENoYW5naW5nIHRoZSAkc2VsZWN0IHF1ZXJ5IHdpbGwgcmVzdWx0IGluIGEgY2xpZW50IHNpZGUgY2FjaGUgbWlzcyBhbmQgdGhlIG5ldyB2YWx1ZSB3aWxsIGJlIHJldHVybmVkXHJcbiAgICAgIHZhciBxdWVyeTQgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFxyXG4gICAgICAgIFwiYWNjb3VudFwiLFxyXG4gICAgICAgIGFjY291bnRpZC5pZCxcclxuICAgICAgICBcIj8kc2VsZWN0PW5hbWUmJGV4cGFuZD1wcmltYXJ5Y29udGFjdGlkKCRzZWxlY3Q9bGFzdG5hbWUpXCJcclxuICAgICAgKTtcclxuICAgICAgdmFyIGV0YWc0ID0gcXVlcnk0W1wiQG9kYXRhLmV0YWdcIl07XHJcbiAgIFxyXG4gICAgICBhc3NlcnQuZXF1YWwoZXRhZzEsIGV0YWc0LCBcIlJlY29yZCBub3QgbW9kaWZpZWRcIik7XHJcbiAgICAgIGFzc2VydC5lcXVhbChcclxuICAgICAgICBcIlNhbXBsZSBDb250YWN0IChlZGl0ZWQpXCIsXHJcbiAgICAgICAgcXVlcnk0LnByaW1hcnljb250YWN0aWQubGFzdG5hbWUsXHJcbiAgICAgICAgXCJDb250YWN0IGNoYW5nZWRcIlxyXG4gICAgICApO1xyXG4gICAgfSBmaW5hbGx5IHtcclxuICAgICAgLy8gRGVsZXRlIGFjY291bnQgJiBjb250YWN0XHJcbiAgICAgIGlmIChhY2NvdW50aWQuaWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudGlkLmlkKTtcclxuICAgICAgfVxyXG4gICAgICBpZiAoY29udGFjdGlkLmlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJjb250YWN0XCIsIGNvbnRhY3RpZC5pZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZTpcclxuLy8gIENyZWF0aW5nIGEgcmVjb3JkIGFuZCB0aGVuIHJldHJpZXZpbmcgaXQgdW5jaGFuZ2VkIHdpdGggdGhlIHNhbWUgZXRhZ1xyXG4vLyAgVGhpcyBpcyBhIHRlY2huaXF1ZSBlbXBsb3llZCBieSB0aGUgY2xpZW50IHNpZGUgYXBpIHRvIGF2b2lkIHVuZWNjZXNhcnkgZGF0YSBiZWluZyB0cmFuc2ZlcnJlZCBvdmVyIHRoZSBuZXR3b3JrXHJcbi8vICBTZWU6IGh0dHBzOi8vZG9jcy5taWNyb3NvZnQuY29tL2VuLXVzL2R5bmFtaWNzMzY1L2N1c3RvbWVyLWVuZ2FnZW1lbnQvZGV2ZWxvcGVyL3dlYmFwaS9jcmVhdGUtZW50aXR5LXdlYi1hcGlcclxuXHJcbmRlc2NyaWJlKFwiXCIsIGZ1bmN0aW9uKCkge1xyXG4gIGl0KFwiZXRhZ3NcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG5cclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuICAgIFxyXG4gICAgdmFyIGFjY291bnQ6IHtcclxuICAgICAgYWNjb3VudGlkPzogc3RyaW5nO1xyXG4gICAgICBuYW1lPzogc3RyaW5nO1xyXG4gICAgICBjcmVkaXRsaW1pdDogTnVtYmVyOyAvLyBNb25leVxyXG4gICAgICBbaW5kZXg6IHN0cmluZ106IGFueTsgLy8gQWxsb3cgc2V0dGluZyBAb2RhdGEgcHJvcGVydGllc1xyXG4gICAgfTtcclxuXHJcbiAgICBhY2NvdW50ID0ge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBBY2NvdW50XCIsXHJcbiAgICAgIGNyZWRpdGxpbWl0OiAxMDAwLFxyXG4gICAgfTtcclxuXHJcbiAgICB0cnkge1xyXG4gICAgICAvLyBDcmVhdGUgQWNjb3VudFxyXG4gICAgICBhY2NvdW50LmFjY291bnRpZCA9IChhd2FpdCAoPGFueT4oXHJcbiAgICAgICAgWHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIGFjY291bnQpXHJcbiAgICAgICkpKS5pZDtcclxuXHJcbiAgICAgIGlmICghYWNjb3VudC5hY2NvdW50aWQpXHJcbiAgICAgICAgdGhyb3cgbmV3IEVycm9yKFwiQWNjb3VudCBub3QgY3JlYXRlZFwiKTtcclxuXHJcbiAgICAgIC8vIFJlYWQgdGhlIEFjY291bnRcclxuICAgICAgdmFyIHF1ZXJ5MSA9IGF3YWl0IFhybS5XZWJBcGkucmV0cmlldmVSZWNvcmQoXCJhY2NvdW50XCIsYWNjb3VudC5hY2NvdW50aWQsXCI/JHNlbGVjdD1uYW1lXCIpO1xyXG4gICAgICB2YXIgZXRhZzEgPSBxdWVyeTFbXCJAb2RhdGEuZXRhZ1wiXTtcclxuXHJcbiAgICAgIC8vIFJlYWQgdGhlIEFjY291bnQgYWdhaW4gKFRoaXMgd2lsbCByZXR1cm4gdGhlIHNhbWUgcmVjb3JkIGFzIGFib3ZlIHNpbmNlIHRoZSBzZXJ2ZXIgd2lsbCByZXR1cm4gMzA0IE5vdCBNb2RpZmllZCApXHJcbiAgICAgIHZhciBxdWVyeTIgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFwiYWNjb3VudFwiLGFjY291bnQuYWNjb3VudGlkLFwiPyRzZWxlY3Q9bmFtZVwiKTtcclxuICAgICAgdmFyIGV0YWcyID0gcXVlcnkyW1wiQG9kYXRhLmV0YWdcIl07XHJcblxyXG4gICAgICBhc3NlcnQuZXF1YWwoZXRhZzEsZXRhZzIsXCJSZWNvcmQgbm90IG1vZGlmaWVkXCIpO1xyXG5cclxuICAgICAgLy8gVXBkYXRlIHRoZSB2YWx1ZVxyXG4gICAgICBhY2NvdW50Lm5hbWUgPSBcIlNhbXBsZSBBY2NvdW50ICh1cGRhdGVkKVwiO1xyXG4gICAgICBhd2FpdCBYcm0uV2ViQXBpLnVwZGF0ZVJlY29yZChcImFjY291bnRcIixhY2NvdW50LmFjY291bnRpZCwgYWNjb3VudCk7XHJcblxyXG4gICAgICAvLyBSZWFkIHRoZSBBY2NvdW50IGFnYWluLiBTaW5jZSB0aGUgcmVjb3JkIGlzIHVwZGF0ZWQgb24gdGhlIHNlcnZlciBpdCB3aWxsIGhhdmUgYSBkaWZmZXJlbnQgZXRhZ1xyXG4gICAgICB2YXIgcXVlcnkzID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZVJlY29yZChcImFjY291bnRcIixhY2NvdW50LmFjY291bnRpZCxcIj8kc2VsZWN0PW5hbWVcIik7XHJcbiAgICAgIHZhciBldGFnMyA9IHF1ZXJ5M1tcIkBvZGF0YS5ldGFnXCJdO1xyXG5cclxuICAgICAgYXNzZXJ0Lm5vdEVxdWFsKGV0YWcxLGV0YWczLFwiUmVjb3JkIG1vZGlmaWVkXCIpO1xyXG5cclxuXHJcbiAgICB9IGZpbmFsbHkge1xyXG4gICAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgICBpZiAoYWNjb3VudC5hY2NvdW50aWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudC5hY2NvdW50aWQpO1xyXG4gICAgICB9XHJcbiAgICB9XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWU6XHJcbi8vICBRdWVyeWluZyByZWNvcmRzIHVzaW5nIGZldGNoeG1sXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIlF1ZXJ5IHdpdGggRmV0Y2hYbWxcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG5cclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuICAgIC8vIENoZWNrIHRoZSBhY2NvdW50IGhhcyBiZWVuIGNyZWF0ZWRcclxuICAgIHZhciBmZXRjaCA9IGA8ZmV0Y2ggbm8tbG9jaz1cInRydWVcIiA+XHJcbiAgICAgIDxlbnRpdHkgbmFtZT1cImFjY291bnRcIj5cclxuICAgICAgICA8YXR0cmlidXRlIG5hbWU9XCJuYW1lXCIvPlxyXG4gICAgICA8L2VudGl0eT5cclxuICAgIDwvZmV0Y2g+YDtcclxuXHJcbiAgICB2YXIgYWNjb3VudHMgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlTXVsdGlwbGVSZWNvcmRzKFxyXG4gICAgICBcImFjY291bnRcIixcclxuICAgICAgXCI/ZmV0Y2hYbWw9XCIgKyBmZXRjaFxyXG4gICAgKTtcclxuXHJcbiAgICBhc3NlcnQuaXNOb3ROdWxsKGFjY291bnRzLmVudGl0aWVzLCBcIkFjY291bnQgcXVlcnkgcmV0dXJucyByZXN1bHRzXCIpO1xyXG4gIH0pO1xyXG59KTtcclxuIiwiLy8gRGVtb25zdHJhdGVzIHRoZSBmb2xsb3dpbmcgdGVjaG5pcXVlOlxyXG4vLyAgUXVlcnlpbmcgZm9yIGEgcmVjb3JkIGJ5IGlkIHVzaW5nIHJldHJpZXZlUmVjb3JkIHdpdGggYSAkc2VsZWN0IGNsYXVzZVxyXG5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJSZWFkXCIsIGFzeW5jIGZ1bmN0aW9uKCkge1xyXG4gICAgdGhpcy50aW1lb3V0KDkwMDAwKTtcclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuXHJcbiAgICB2YXIgcmVjb3JkOiB7XHJcbiAgICAgIGFjY291bnRpZD86IHN0cmluZztcclxuICAgICAgbmFtZT86IHN0cmluZztcclxuICAgIH07XHJcblxyXG4gICAgcmVjb3JkID0gIHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiXHJcbiAgICB9O1xyXG5cclxuICAgIC8vIENyZWF0ZSBBY2NvdW50XHJcbiAgICByZWNvcmQuYWNjb3VudGlkID0gKGF3YWl0ICg8YW55PihcclxuICAgICAgWHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIHJlY29yZClcclxuICAgICkpKS5pZDtcclxuXHJcbiAgICBpZiAoIXJlY29yZC5hY2NvdW50aWQpXHJcbiAgICAgIHRocm93IG5ldyBFcnJvcihcIkFjY291bnQgbm90IGNyZWF0ZWRcIik7XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgdmFyIGFjY291bnRzUmVhZCA9IGF3YWl0IFhybS5XZWJBcGkucmV0cmlldmVSZWNvcmQoXHJcbiAgICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgICAgcmVjb3JkLmFjY291bnRpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PW5hbWUscHJpbWFyeWNvbnRhY3RpZFwiXHJcbiAgICAgICk7XHJcblxyXG4gICAgICBpZiAoIWFjY291bnRzUmVhZCB8fCAhYWNjb3VudHNSZWFkLm5hbWUpIHtcclxuICAgICAgICB0aHJvdyBuZXcgRXJyb3IoXCJBY2NvdW50IG5vdCBjcmVhdGVkXCIpO1xyXG4gICAgICB9XHJcbiAgICAgIGFzc2VydC5lcXVhbChhY2NvdW50c1JlYWQubmFtZSwgcmVjb3JkLm5hbWUsIFwiQWNjb3VudCBjcmVhdGVkXCIpO1xyXG5cclxuICAgIH0gZmluYWxseSB7XHJcbiAgICAgIC8vIERlbGV0ZSBhY2NvdW50XHJcbiAgICAgIGlmIChyZWNvcmQuYWNjb3VudGlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJhY2NvdW50XCIsIHJlY29yZC5hY2NvdW50aWQpO1xyXG4gICAgICB9XHJcbiAgICB9XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWU6XHJcbi8vICBRdWVyeWluZyBmb3IgbXVsdGlwbGUgcmVjb3JkcyB1c2luZyBvZGF0YSBxdWVyeVxyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvcXVlcnktZGF0YS13ZWItYXBpXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIlJldHJpZXZlTXVsdGlwbGVcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB2YXIgYXNzZXJ0ID0gY2hhaS5hc3NlcnQ7XHJcbiAgICB2YXIgYWNjb3VudGlkOiB7XHJcbiAgICAgIGVudGl0eVR5cGU6IHN0cmluZztcclxuICAgICAgaWQ6IHN0cmluZztcclxuICAgIH0gPSBhd2FpdCAoPGFueT5Ycm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImFjY291bnRcIiwge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBBY2NvdW50XCIsXHJcbiAgICAgIHJldmVudWU6IDIwMDAwLjAxXHJcbiAgICB9KSk7XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgdmFyIHJlc3VsdHMgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlTXVsdGlwbGVSZWNvcmRzKFxyXG4gICAgICAgIFwiYWNjb3VudFwiLFxyXG4gICAgICAgIFwiPyRzZWxlY3Q9bmFtZSYkZmlsdGVyPXJldmVudWUgZ3QgMjAwMDAgYW5kIHJldmVudWUgbHQgMjAwMDEgYW5kIG5hbWUgZXEgJ1NhbXBsZSBBY2NvdW50J1wiLFxyXG4gICAgICAgIDEwXHJcbiAgICAgICk7XHJcblxyXG4gICAgICAvLyBDaGVjayB0aGF0IHRoZXJlIGlzIGEgc2luZ2xlIHJlc3VsdCByZXR1cm5lZFxyXG4gICAgICBpZiAoIXJlc3VsdHMuZW50aXRpZXMgfHwgIXJlc3VsdHMuZW50aXRpZXMubGVuZ3RoKVxyXG4gICAgICAgIHRocm93IG5ldyBFcnJvcihcIk5vIHJlc3VsdHMgcmV0dXJuZWRcIik7XHJcblxyXG4gICAgICBhc3NlcnQuZXF1YWwocmVzdWx0cy5lbnRpdGllcy5sZW5ndGgsIDEsIFwiU2luZ2xlIHJlc3VsdCByZXR1cm5lZFwiKTtcclxuICAgIH0gZmluYWxseSB7XHJcbiAgICAgIC8vIERlbGV0ZSBhY2NvdW50XHJcbiAgICAgIGlmIChhY2NvdW50aWQuaWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudGlkLmlkKTtcclxuICAgICAgfVxyXG4gICAgfVxyXG4gIH0pO1xyXG59KTtcclxuIiwiLy8gRGVtb25zdHJhdGVzIHRoZSBmb2xsb3dpbmcgdGVjaG5pcXVlOlxyXG4vLyAgVXBkYXRpbmcgYSByZWNvcmQncyBTdGF0ZSBhbmQgU3RhdHVzIFJlc2FvblxyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvY3JlYXRlLWVudGl0eS13ZWItYXBpXHJcbmNvbnN0IGVudW0gb3Bwb3J0dW5pdHlfc3RhdHVzY29kZSB7XHJcbiAgSW5fUHJvZ3Jlc3MgPSAxLFxyXG4gIE9uX0hvbGQgPSAyLFxyXG4gIFdvbiA9IDMsXHJcbiAgQ2FuY2VsZWQgPSA0LFxyXG4gIE91dFNvbGQgPSA1XHJcbn1cclxuY29uc3QgZW51bSBhY2NvdW50X3N0YXRlY29kZSB7XHJcbiAgQWN0aXZlID0gMCxcclxuICBJbmFjdGl2ZSA9IDFcclxufVxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIlNldFN0YXRlXCIsIGFzeW5jIGZ1bmN0aW9uKCkge1xyXG4gICAgdGhpcy50aW1lb3V0KDkwMDAwKTtcclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuXHJcbiAgICB2YXIgYWNjb3VudDoge1xyXG4gICAgICBhY2NvdW50aWQ/OiBzdHJpbmc7XHJcbiAgICAgIG5hbWU/OiBzdHJpbmc7XHJcbiAgICAgIGFkZHJlc3MxX2NpdHk/OiBzdHJpbmc7XHJcbiAgICAgIHN0YXRlY29kZT86IGFjY291bnRfc3RhdGVjb2RlO1xyXG4gICAgfTtcclxuXHJcbiAgICBhY2NvdW50ID0ge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBBY2NvdW50XCJcclxuICAgIH07XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgLy8gQ3JlYXRlIEFjY291bnRcclxuICAgICAgYWNjb3VudC5hY2NvdW50aWQgPSAoYXdhaXQgKDxhbnk+KFxyXG4gICAgICAgIFhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiYWNjb3VudFwiLCBhY2NvdW50KVxyXG4gICAgICApKSkuaWQ7XHJcblxyXG4gICAgICBpZiAoIWFjY291bnQuYWNjb3VudGlkKSB7XHJcbiAgICAgICAgdGhyb3cgbmV3IEVycm9yKFwiQWNjb3VudCBJRCBub3QgcmV0dXJuZWRcIik7XHJcbiAgICAgIH1cclxuXHJcbiAgICAgIC8vIENyZWF0ZSBPcHBvcnR1bml0eSBmb3IgdGhlIGFjY291bnRcclxuICAgICAgdmFyIG9wcG9ydHVuaXR5OiBhbnkgPSB7XHJcbiAgICAgICAgbmFtZTogXCJTYW1wbGUgT3Bwb3J0dW5pdHlcIixcclxuICAgICAgICBcInBhcmVudGFjY291bnRpZEBvZGF0YS5iaW5kXCI6IGBhY2NvdW50cygke2FjY291bnQuYWNjb3VudGlkfSlgXHJcbiAgICAgIH07XHJcblxyXG4gICAgICB2YXIgY3JlYXRlT3Bwb3J0dW5pdHlSZXNwb25zZToge1xyXG4gICAgICAgIGVudGl0eVR5cGU6IFN0cmluZztcclxuICAgICAgICBpZDogU3RyaW5nO1xyXG4gICAgICB9ID0gPGFueT5hd2FpdCBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcIm9wcG9ydHVuaXR5XCIsIG9wcG9ydHVuaXR5KTtcclxuXHJcbiAgICAgIG9wcG9ydHVuaXR5Lm9wcG9ydHVuaXR5aWQgPSBjcmVhdGVPcHBvcnR1bml0eVJlc3BvbnNlLmlkO1xyXG5cclxuICAgICAgLy8gQ2hhbmdlIE9wcG9ydHVuaXR5IFN0YXR1cyBSZWFzb24gdG8gSW4gUHJvZ3Jlc3NcclxuICAgICAgb3Bwb3J0dW5pdHkuc3RhdHVzY29kZSA9IG9wcG9ydHVuaXR5X3N0YXR1c2NvZGUuSW5fUHJvZ3Jlc3M7XHJcbiAgICAgIGF3YWl0IFhybS5XZWJBcGkudXBkYXRlUmVjb3JkKFxyXG4gICAgICAgIFwib3Bwb3J0dW5pdHlcIixcclxuICAgICAgICBvcHBvcnR1bml0eS5vcHBvcnR1bml0eWlkLFxyXG4gICAgICAgIG9wcG9ydHVuaXR5XHJcbiAgICAgICk7XHJcblxyXG4gICAgICAvLyBDaGVjayB0aGUgb3Bwb3J0dW5pdHkgaXMgdXBkYXRlZFxyXG4gICAgICB2YXIgb3Bwb3J0dW5pdHlSZWFkID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZVJlY29yZChcclxuICAgICAgICBcIm9wcG9ydHVuaXR5XCIsXHJcbiAgICAgICAgb3Bwb3J0dW5pdHkub3Bwb3J0dW5pdHlpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PXN0YXR1c2NvZGVcIlxyXG4gICAgICApO1xyXG5cclxuICAgICAgaWYgKCFvcHBvcnR1bml0eVJlYWQgfHwgIW9wcG9ydHVuaXR5UmVhZC5zdGF0dXNjb2RlKSB7XHJcbiAgICAgICAgdGhyb3cgbmV3IEVycm9yKFwiT3Bwb3J0dW5pdHkgbm90IHVwZGF0ZWRcIik7XHJcbiAgICAgIH1cclxuICAgICAgYXNzZXJ0LmVxdWFsKFxyXG4gICAgICAgIG9wcG9ydHVuaXR5UmVhZC5zdGF0dXNjb2RlLFxyXG4gICAgICAgIG9wcG9ydHVuaXR5X3N0YXR1c2NvZGUuSW5fUHJvZ3Jlc3MsXHJcbiAgICAgICAgXCJPcHBvcnR1bml0eSBJbiBQcm9ncmVzc1wiXHJcbiAgICAgICk7XHJcblxyXG4gICAgICAvLyBVcGRhdGUgYWNjb3VudCBzdGF0ZSB0byBJbiBBY3RpdmVcclxuICAgICAgYWNjb3VudC5zdGF0ZWNvZGUgPSBhY2NvdW50X3N0YXRlY29kZS5JbmFjdGl2ZTtcclxuICAgICAgYXdhaXQgWHJtLldlYkFwaS51cGRhdGVSZWNvcmQoXCJhY2NvdW50XCIsIGFjY291bnQuYWNjb3VudGlkLCBhY2NvdW50KTtcclxuXHJcbiAgICAgIC8vIENoZWNrIHRoZSBhY2NvdW50IGlzIHVwZGF0ZWRcclxuICAgICAgdmFyIGFjY291bnRSZWFkID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZVJlY29yZChcclxuICAgICAgICBcImFjY291bnRcIixcclxuICAgICAgICBhY2NvdW50LmFjY291bnRpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PXN0YXRlY29kZVwiXHJcbiAgICAgICk7XHJcblxyXG4gICAgICBpZiAoIWFjY291bnRSZWFkIHx8IGFjY291bnRSZWFkLnN0YXRlY29kZSA9PSB1bmRlZmluZWQpIHtcclxuICAgICAgICB0aHJvdyBuZXcgRXJyb3IoXCJBY2NvdW50IG5vdCB1cGRhdGVkXCIpO1xyXG4gICAgICB9XHJcbiAgICAgIGFzc2VydC5lcXVhbChcclxuICAgICAgICBhY2NvdW50UmVhZC5zdGF0ZWNvZGUsXHJcbiAgICAgICAgYWNjb3VudF9zdGF0ZWNvZGUuSW5hY3RpdmUsXHJcbiAgICAgICAgXCJBY2NvdW50IEluYWN0aXZlXCJcclxuICAgICAgKTtcclxuICAgIH0gZmluYWxseSB7XHJcbiAgICAgIC8vIERlbGV0ZSBhY2NvdW50XHJcbiAgICAgIGlmIChhY2NvdW50LmFjY291bnRpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiYWNjb3VudFwiLCBhY2NvdW50LmFjY291bnRpZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZTpcclxuLy8gIFVwZGF0aW5nIGEgcmVjb3JkXHJcbi8vICBTZWU6IGh0dHBzOi8vZG9jcy5taWNyb3NvZnQuY29tL2VuLXVzL2R5bmFtaWNzMzY1L2N1c3RvbWVyLWVuZ2FnZW1lbnQvZGV2ZWxvcGVyL3dlYmFwaS9jcmVhdGUtZW50aXR5LXdlYi1hcGlcclxuXHJcbmRlc2NyaWJlKFwiXCIsIGZ1bmN0aW9uKCkge1xyXG4gIGl0KFwiVXBkYXRlXCIsIGFzeW5jIGZ1bmN0aW9uKCkge1xyXG4gICAgdGhpcy50aW1lb3V0KDkwMDAwKTtcclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuXHJcbiAgICB2YXIgcmVjb3JkOiB7XHJcbiAgICAgIGFjY291bnRpZD86IHN0cmluZztcclxuICAgICAgbmFtZT86IHN0cmluZztcclxuICAgICAgYWRkcmVzczFfY2l0eT86IHN0cmluZztcclxuICAgIH07XHJcbiAgICBcclxuICAgIHJlY29yZCA9ICB7XHJcbiAgICAgIG5hbWU6IFwiU2FtcGxlIEFjY291bnRcIlxyXG4gICAgfTtcclxuXHJcbiAgICB0cnkge1xyXG4gICAgICAvLyBDcmVhdGUgQWNjb3VudFxyXG4gICAgICByZWNvcmQuYWNjb3VudGlkID0gKGF3YWl0ICg8YW55PihcclxuICAgICAgICBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImFjY291bnRcIiwgcmVjb3JkKVxyXG4gICAgICApKSkuaWQ7XHJcblxyXG4gICAgICBpZiAoIXJlY29yZC5hY2NvdW50aWQpXHJcbiAgICAgIHtcclxuICAgICAgICB0aHJvdyBuZXcgRXJyb3IoXCJBY2NvdW50IElEIG5vdCByZXR1cm5lZFwiKTtcclxuICAgICAgfVxyXG4gICAgICByZWNvcmQubmFtZSA9IFwiU2FtcGxlIEFjY291bnQgKHVwZGF0ZWQpXCI7XHJcbiAgICAgIHJlY29yZC5hZGRyZXNzMV9jaXR5ID0gXCJPeGZvcmRcIjtcclxuXHJcbiAgICAgIC8vIFVwZGF0ZSBhY2NvdW50XHJcbiAgICAgIGF3YWl0IFhybS5XZWJBcGkudXBkYXRlUmVjb3JkKFwiYWNjb3VudFwiLHJlY29yZC5hY2NvdW50aWQsIHJlY29yZCk7XHJcblxyXG4gICAgICAvLyBDaGVjayB0aGUgcmVjb3JkIGlzIHVwZGF0ZWRcclxuICAgICAgdmFyIGFjY291bnRzUmVhZCA9IGF3YWl0IFhybS5XZWJBcGkucmV0cmlldmVSZWNvcmQoXHJcbiAgICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgICAgcmVjb3JkLmFjY291bnRpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PW5hbWUsYWRkcmVzczFfY2l0eVwiXHJcbiAgICAgICk7XHJcbiAgICAgIFxyXG4gICAgICBpZiAoIWFjY291bnRzUmVhZCB8fCAhYWNjb3VudHNSZWFkLm5hbWUgfHwgIWFjY291bnRzUmVhZC5hZGRyZXNzMV9jaXR5KSB7XHJcbiAgICAgICAgdGhyb3cgbmV3IEVycm9yKFwiQWNjb3VudCBub3QgdXBkYXRlZFwiKTtcclxuICAgICAgfVxyXG4gICAgICBhc3NlcnQuZXF1YWwoYWNjb3VudHNSZWFkLm5hbWUsIHJlY29yZC5uYW1lLCBcIkFjY291bnQgdXBkYXRlZFwiKTtcclxuICAgICAgYXNzZXJ0LmVxdWFsKGFjY291bnRzUmVhZC5hZGRyZXNzMV9jaXR5LCByZWNvcmQuYWRkcmVzczFfY2l0eSwgXCJBY2NvdW50IHVwZGF0ZWRcIik7XHJcblxyXG4gICAgfSBmaW5hbGx5IHtcclxuICAgICAgLy8gRGVsZXRlIGFjY291bnRcclxuICAgICAgaWYgKHJlY29yZC5hY2NvdW50aWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgcmVjb3JkLmFjY291bnRpZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZXM6XHJcbi8vICBVc2luZyB0aGUgQ2FsY3VsYXRlUm9sbHVwIFVuYm91bmQgRnVuY3Rpb24gdG8gcmVjYWxjdWxhdGUgYSByb2xsdXAgZmllbGRcclxuLy8gIFNlZTogaHR0cHM6Ly9kb2NzLm1pY3Jvc29mdC5jb20vZW4tdXMvZHluYW1pY3MzNjUvY3VzdG9tZXItZW5nYWdlbWVudC9kZXZlbG9wZXIvd2ViYXBpL3VzZS13ZWItYXBpLWZ1bmN0aW9uc1xyXG5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJDYWxjdWxhdGVSb2xsdXBcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG5cclxuXHJcbiAgICAvLyBHZXQgYW4gYWNjb3VudFxyXG4gICAgdmFyIHJlc3BvbnNlOiB7XHJcbiAgICAgIGVudGl0aWVzOiB7XHJcbiAgICAgICAgYWNjb3VudGlkOiBTdHJpbmc7XHJcbiAgICAgICAgbmFtZTogU3RyaW5nO1xyXG4gICAgICB9W107XHJcbiAgICAgIG5leHRMaW5rOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZU11bHRpcGxlUmVjb3JkcyhcclxuICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgIFwiPyRzZWxlY3Q9YWNjb3VudGlkLG5hbWUmJHRvcD0xXCIsXHJcbiAgICAgIDFcclxuICAgICk7XHJcblxyXG4gICAgLy8gRXhlY3V0ZSByZXF1ZXN0XHJcbiAgICB2YXIgcmVxdWVzdCA9IG5ldyBjbGFzcyB7XHJcbiAgICAgIGdldE1ldGFkYXRhKCk6IGFueSB7XHJcbiAgICAgICAgcmV0dXJuIHtcclxuICAgICAgICAgIHBhcmFtZXRlclR5cGVzOiB7XHJcbiAgICAgICAgICAgIEZpZWxkTmFtZToge1xyXG4gICAgICAgICAgICAgIHR5cGVOYW1lOiBcIkVkbS5TdHJpbmdcIixcclxuICAgICAgICAgICAgICBzdHJ1Y3R1cmFsUHJvcGVydHk6IDFcclxuICAgICAgICAgICAgfSxcclxuICAgICAgICAgICAgVGFyZ2V0OiB7XHJcbiAgICAgICAgICAgICAgdHlwZU5hbWU6IFwibXNjcm0uY3JtYmFzZWVudGl0eVwiLFxyXG4gICAgICAgICAgICAgIHN0cnVjdHVyYWxQcm9wZXJ0eTogNVxyXG4gICAgICAgICAgICB9XHJcbiAgICAgICAgICB9LFxyXG4gICAgICAgICAgb3BlcmF0aW9uVHlwZTogMSxcclxuICAgICAgICAgIG9wZXJhdGlvbk5hbWU6IFwiQ2FsY3VsYXRlUm9sbHVwRmllbGRcIlxyXG4gICAgICAgIH07XHJcbiAgICAgIH1cclxuICAgICAgVGFyZ2V0ID0ge1xyXG4gICAgICAgIGlkOiByZXNwb25zZS5lbnRpdGllc1swXS5hY2NvdW50aWQsXHJcbiAgICAgICAgZW50aXR5VHlwZTogXCJhY2NvdW50XCJcclxuICAgICAgfTtcclxuICAgICAgRmllbGROYW1lID0gXCJvcGVuZGVhbHNcIjtcclxuICAgIH0oKTtcclxuICAgIGF3YWl0IFhybS5XZWJBcGkub25saW5lLmV4ZWN1dGUocmVxdWVzdCk7XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWVzOlxyXG4vLyAgVXNpbmcgdGhlIENhbGN1bGF0ZVRvdGFsVGltZUluY2lkZW50IEJvdW5kIEZ1bmN0aW9uIHRvIHJlY2FsY3VsYXRlIGEgcm9sbHVwIGZpZWxkXHJcbi8vICBTZWU6IGh0dHBzOi8vZG9jcy5taWNyb3NvZnQuY29tL2VuLXVzL2R5bmFtaWNzMzY1L2N1c3RvbWVyLWVuZ2FnZW1lbnQvZGV2ZWxvcGVyL3dlYmFwaS91c2Utd2ViLWFwaS1mdW5jdGlvbnNcclxuXHJcbmRlc2NyaWJlKFwiXCIsIGZ1bmN0aW9uKCkge1xyXG4gIGl0KFwiQ2FsY3VsYXRlVG90YWxUaW1lSW5jaWRlbnRcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG5cclxuICAgIC8vIEdldCBhbiBpbmNpZGVudFxyXG4gICAgdmFyIHJlc3BvbnNlOiB7XHJcbiAgICAgIGVudGl0aWVzOiB7XHJcbiAgICAgICAgaW5jaWRlbnRpZDogU3RyaW5nO1xyXG4gICAgICB9W107XHJcbiAgICAgIG5leHRMaW5rOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgWHJtLldlYkFwaS5yZXRyaWV2ZU11bHRpcGxlUmVjb3JkcyhcclxuICAgICAgXCJpbmNpZGVudFwiLFxyXG4gICAgICBcIj8kc2VsZWN0PWluY2lkZW50aWRcIixcclxuICAgICAgMVxyXG4gICAgKTtcclxuXHJcbiAgICAvLyBFeGVjdXRlIENhbGN1bGF0ZVRvdGFsVGltZUluY2lkZW50IHJlcXVlc3RcclxuICAgIC8vIFRoaXMgaXMgYSBib3VuZCBmdW5jdGlvbiB3aGljaCB3ZSBwYXNzIHRoZSBlbnRpdHkgcGFyYW1ldGVyIGFzIHRoZSB0YXJnZXQgaW5jaWRlbnRcclxuICAgIHZhciByZXF1ZXN0ID0gbmV3IGNsYXNzIHtcclxuICAgICAgZW50aXR5ID0ge1xyXG4gICAgICAgIGlkOiByZXNwb25zZS5lbnRpdGllc1swXS5pbmNpZGVudGlkLFxyXG4gICAgICAgIGVudGl0eVR5cGU6IFwiaW5jaWRlbnRcIlxyXG4gICAgICB9O1xyXG5cclxuICAgICAgZ2V0TWV0YWRhdGEoKTogYW55IHtcclxuICAgICAgICByZXR1cm4ge1xyXG4gICAgICAgICAgYm91bmRQYXJhbWV0ZXI6IFwiZW50aXR5XCIsXHJcbiAgICAgICAgICBwYXJhbWV0ZXJUeXBlczoge1xyXG4gICAgICAgICAgICBlbnRpdHk6IHtcclxuICAgICAgICAgICAgICB0eXBlTmFtZTogXCJtc2NybS5pbmNpZGVudFwiLFxyXG4gICAgICAgICAgICAgIHN0cnVjdHVyYWxQcm9wZXJ0eTogNVxyXG4gICAgICAgICAgICB9XHJcbiAgICAgICAgICB9LFxyXG4gICAgICAgICAgb3BlcmF0aW9uVHlwZTogMSxcclxuICAgICAgICAgIG9wZXJhdGlvbk5hbWU6IFwiQ2FsY3VsYXRlVG90YWxUaW1lSW5jaWRlbnRcIlxyXG4gICAgICAgIH07XHJcbiAgICAgIH1cclxuICAgIH0oKTtcclxuXHJcbiAgICAvLyBUaGUganNvbiBmdW5jdGlvbiBpcyBhIHByb21pc2Ugd2hpY2ggcmV0dXJucyB0aGUgcmVzcG9uc2Ugb2JqZWN0LlxyXG4gICAgdmFyIHJhd1Jlc3BvbnNlID0gYXdhaXQgKDxhbnk+YXdhaXQgWHJtLldlYkFwaS5vbmxpbmUuZXhlY3V0ZShyZXF1ZXN0KSkuanNvbigpOztcclxuICAgXHJcbiAgICBhc3NlcnQuaXNOdW1iZXIocmF3UmVzcG9uc2UuVG90YWxUaW1lLFwiVG90YWwgVGltZSByZXR1cm5lZFwiKTtcclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZXM6XHJcbi8vICBRdWVyeWluZyBtZXRhZGF0YSB1c2luZyB0aGUgUmV0cmlldmVNZXRhZGF0YUNoYW5nZXMgcmVxdWVzdFxyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvdXNlLXdlYi1hcGktZnVuY3Rpb25zXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIlJldHJpZXZlTWV0YWRhdGFDaGFuZ2VzXCIsIGFzeW5jIGZ1bmN0aW9uKCkge1xyXG4gICAgdGhpcy50aW1lb3V0KDkwMDAwKTtcclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuXHJcbiAgICAvLyBRdWVyeSBmb3IgQWNjb3VudCBtZXRhZGF0YSBhbmQgcmV0dXJuIG9ubHkgdGhlIE9wdGlvblNldCB2YWx1ZXMgZm9yIHRoZSBhdHRyaWJ1dGUgYWRkcmVzczFfc2hpcHBpbmdtZXRob2Rjb2RlXHJcbiAgICB2YXIgcmVxdWVzdCA9IG5ldyBjbGFzcyB7XHJcbiAgICAgIFF1ZXJ5ID0ge1xyXG4gICAgICAgIENyaXRlcmlhOiB7XHJcbiAgICAgICAgICBDb25kaXRpb25zOiBbXHJcbiAgICAgICAgICAgIHtcclxuICAgICAgICAgICAgICBQcm9wZXJ0eU5hbWU6IFwiTG9naWNhbE5hbWVcIixcclxuICAgICAgICAgICAgICBDb25kaXRpb25PcGVyYXRvcjogXCJFcXVhbHNcIixcclxuICAgICAgICAgICAgICBWYWx1ZToge1xyXG4gICAgICAgICAgICAgICAgVmFsdWU6IFwiYWNjb3VudFwiLFxyXG4gICAgICAgICAgICAgICAgVHlwZTogXCJTeXN0ZW0uU3RyaW5nXCJcclxuICAgICAgICAgICAgICB9XHJcbiAgICAgICAgICAgIH1cclxuICAgICAgICAgIF0sXHJcbiAgICAgICAgICBGaWx0ZXJPcGVyYXRvcjogXCJBbmRcIlxyXG4gICAgICAgIH0sXHJcbiAgICAgICAgUHJvcGVydGllczoge1xyXG4gICAgICAgICAgUHJvcGVydHlOYW1lczogW1wiQXR0cmlidXRlc1wiXVxyXG4gICAgICAgIH0sXHJcbiAgICAgICAgQXR0cmlidXRlUXVlcnk6IHtcclxuICAgICAgICAgIFByb3BlcnRpZXM6IHtcclxuICAgICAgICAgICAgUHJvcGVydHlOYW1lczogW1wiT3B0aW9uU2V0XCJdXHJcbiAgICAgICAgICB9LFxyXG4gICAgICAgICAgQ3JpdGVyaWE6IHtcclxuICAgICAgICAgICAgQ29uZGl0aW9uczogW1xyXG4gICAgICAgICAgICAgIHtcclxuICAgICAgICAgICAgICAgIFByb3BlcnR5TmFtZTogXCJMb2dpY2FsTmFtZVwiLFxyXG4gICAgICAgICAgICAgICAgQ29uZGl0aW9uT3BlcmF0b3I6IFwiRXF1YWxzXCIsXHJcbiAgICAgICAgICAgICAgICBWYWx1ZToge1xyXG4gICAgICAgICAgICAgICAgICBWYWx1ZTogXCJhZGRyZXNzMV9zaGlwcGluZ21ldGhvZGNvZGVcIixcclxuICAgICAgICAgICAgICAgICAgVHlwZTogXCJTeXN0ZW0uU3RyaW5nXCJcclxuICAgICAgICAgICAgICAgIH1cclxuICAgICAgICAgICAgICB9XHJcbiAgICAgICAgICAgIF0sXHJcbiAgICAgICAgICAgIEZpbHRlck9wZXJhdG9yOiBcIkFuZFwiXHJcbiAgICAgICAgICB9XHJcbiAgICAgICAgfVxyXG4gICAgICB9O1xyXG5cclxuICAgICAgZ2V0TWV0YWRhdGEoKTogYW55IHtcclxuICAgICAgICByZXR1cm4ge1xyXG4gICAgICAgICAgcGFyYW1ldGVyVHlwZXM6IHtcclxuICAgICAgICAgICAgQXBwTW9kdWxlSWQ6IHtcclxuICAgICAgICAgICAgICB0eXBlTmFtZTogXCJFZG0uR3VpZFwiLFxyXG4gICAgICAgICAgICAgIHN0cnVjdHVyYWxQcm9wZXJ0eTogMVxyXG4gICAgICAgICAgICB9LFxyXG4gICAgICAgICAgICBDbGllbnRWZXJzaW9uU3RhbXA6IHtcclxuICAgICAgICAgICAgICB0eXBlTmFtZTogXCJFZG0uU3RyaW5nXCIsXHJcbiAgICAgICAgICAgICAgc3RydWN0dXJhbFByb3BlcnR5OiAxXHJcbiAgICAgICAgICAgIH0sXHJcbiAgICAgICAgICAgIERlbGV0ZWRNZXRhZGF0YUZpbHRlcnM6IHtcclxuICAgICAgICAgICAgICB0eXBlTmFtZTogXCJtc2NybS5EZWxldGVkTWV0YWRhdGFGaWx0ZXJzXCIsXHJcbiAgICAgICAgICAgICAgc3RydWN0dXJhbFByb3BlcnR5OiAzXHJcbiAgICAgICAgICAgIH0sXHJcbiAgICAgICAgICAgIFF1ZXJ5OiB7XHJcbiAgICAgICAgICAgICAgdHlwZU5hbWU6IFwibXNjcm0uRW50aXR5UXVlcnlFeHByZXNzaW9uXCIsXHJcbiAgICAgICAgICAgICAgc3RydWN0dXJhbFByb3BlcnR5OiA1XHJcbiAgICAgICAgICAgIH1cclxuICAgICAgICAgIH0sXHJcbiAgICAgICAgICBvcGVyYXRpb25UeXBlOiAxLFxyXG4gICAgICAgICAgb3BlcmF0aW9uTmFtZTogXCJSZXRyaWV2ZU1ldGFkYXRhQ2hhbmdlc1wiXHJcbiAgICAgICAgfTtcclxuICAgICAgfVxyXG4gICAgfSgpO1xyXG4gICAgdmFyIHJhd1Jlc3BvbnNlID0gYXdhaXQgWHJtLldlYkFwaS5vbmxpbmUuZXhlY3V0ZShyZXF1ZXN0KTtcclxuICAgIHZhciByZXNwb25zZToge1xyXG4gICAgICBEZWxldGVkTWV0YWRhdGE6IE9iamVjdDtcclxuICAgICAgRW50aXR5TWV0YWRhdGE6IHtcclxuICAgICAgICBBY3Rpdml0eVR5cGVNYXNrOiBhbnk7XHJcbiAgICAgICAgQXR0cmlidXRlczoge1xyXG4gICAgICAgICAgQXR0cmlidXRlT2Y6IGFueTtcclxuICAgICAgICAgIEF0dHJpYnV0ZVR5cGU6IFN0cmluZztcclxuICAgICAgICAgIEF0dHJpYnV0ZVR5cGVOYW1lOiBPYmplY3Q7XHJcbiAgICAgICAgICBBdXRvTnVtYmVyRm9ybWF0OiBhbnk7XHJcbiAgICAgICAgICBDYW5CZVNlY3VyZWRGb3JDcmVhdGU6IGFueTtcclxuICAgICAgICAgIENhbkJlU2VjdXJlZEZvclJlYWQ6IGFueTtcclxuICAgICAgICAgIENhbkJlU2VjdXJlZEZvclVwZGF0ZTogYW55O1xyXG4gICAgICAgICAgQ2FuTW9kaWZ5QWRkaXRpb25hbFNldHRpbmdzOiBhbnk7XHJcbiAgICAgICAgICBDaGlsZFBpY2tsaXN0TG9naWNhbE5hbWVzOiBhbnlbXTtcclxuICAgICAgICAgIENvbHVtbk51bWJlcjogYW55O1xyXG4gICAgICAgICAgRGVmYXVsdEZvcm1WYWx1ZTogYW55O1xyXG4gICAgICAgICAgRGVwcmVjYXRlZFZlcnNpb246IGFueTtcclxuICAgICAgICAgIERlc2NyaXB0aW9uOiBhbnk7XHJcbiAgICAgICAgICBEaXNwbGF5TmFtZTogYW55O1xyXG4gICAgICAgICAgRW50aXR5TG9naWNhbE5hbWU6IGFueTtcclxuICAgICAgICAgIEV4dGVybmFsTmFtZTogYW55O1xyXG4gICAgICAgICAgRm9ybXVsYURlZmluaXRpb246IGFueTtcclxuICAgICAgICAgIEhhc0NoYW5nZWQ6IGFueTtcclxuICAgICAgICAgIEluaGVyaXRzRnJvbTogYW55O1xyXG4gICAgICAgICAgSW50cm9kdWNlZFZlcnNpb246IGFueTtcclxuICAgICAgICAgIElzQXVkaXRFbmFibGVkOiBhbnk7XHJcbiAgICAgICAgICBJc0N1c3RvbUF0dHJpYnV0ZTogYW55O1xyXG4gICAgICAgICAgSXNDdXN0b21pemFibGU6IGFueTtcclxuICAgICAgICAgIElzRGF0YVNvdXJjZVNlY3JldDogYW55O1xyXG4gICAgICAgICAgSXNGaWx0ZXJhYmxlOiBhbnk7XHJcbiAgICAgICAgICBJc0dsb2JhbEZpbHRlckVuYWJsZWQ6IGFueTtcclxuICAgICAgICAgIElzTG9naWNhbDogYW55O1xyXG4gICAgICAgICAgSXNNYW5hZ2VkOiBhbnk7XHJcbiAgICAgICAgICBJc1ByaW1hcnlJZDogYW55O1xyXG4gICAgICAgICAgSXNQcmltYXJ5TmFtZTogYW55O1xyXG4gICAgICAgICAgSXNSZW5hbWVhYmxlOiBhbnk7XHJcbiAgICAgICAgICBJc1JlcXVpcmVkRm9yRm9ybTogYW55O1xyXG4gICAgICAgICAgSXNSZXRyaWV2YWJsZTogYW55O1xyXG4gICAgICAgICAgSXNTZWFyY2hhYmxlOiBhbnk7XHJcbiAgICAgICAgICBJc1NlY3VyZWQ6IGFueTtcclxuICAgICAgICAgIElzU29ydGFibGVFbmFibGVkOiBhbnk7XHJcbiAgICAgICAgICBJc1ZhbGlkRm9yQWR2YW5jZWRGaW5kOiBhbnk7XHJcbiAgICAgICAgICBJc1ZhbGlkRm9yQ3JlYXRlOiBhbnk7XHJcbiAgICAgICAgICBJc1ZhbGlkRm9yRm9ybTogYW55O1xyXG4gICAgICAgICAgSXNWYWxpZEZvckdyaWQ6IGFueTtcclxuICAgICAgICAgIElzVmFsaWRGb3JSZWFkOiBhbnk7XHJcbiAgICAgICAgICBJc1ZhbGlkRm9yVXBkYXRlOiBhbnk7XHJcbiAgICAgICAgICBMaW5rZWRBdHRyaWJ1dGVJZDogYW55O1xyXG4gICAgICAgICAgTG9naWNhbE5hbWU6IFN0cmluZztcclxuICAgICAgICAgIE1ldGFkYXRhSWQ6IFN0cmluZztcclxuICAgICAgICAgIE9wdGlvblNldDogT2JqZWN0O1xyXG4gICAgICAgICAgUGFyZW50UGlja2xpc3RMb2dpY2FsTmFtZTogYW55O1xyXG4gICAgICAgICAgUmVxdWlyZWRMZXZlbDogYW55O1xyXG4gICAgICAgICAgU2NoZW1hTmFtZTogYW55O1xyXG4gICAgICAgICAgU291cmNlVHlwZTogYW55O1xyXG4gICAgICAgICAgU291cmNlVHlwZU1hc2s6IGFueTtcclxuICAgICAgICB9W107XHJcbiAgICAgICAgQXV0b0NyZWF0ZUFjY2Vzc1RlYW1zOiBhbnk7XHJcbiAgICAgICAgQXV0b1JvdXRlVG9Pd25lclF1ZXVlOiBhbnk7XHJcbiAgICAgICAgQ2FuQmVJbkN1c3RvbUVudGl0eUFzc29jaWF0aW9uOiBhbnk7XHJcbiAgICAgICAgQ2FuQmVJbk1hbnlUb01hbnk6IGFueTtcclxuICAgICAgICBDYW5CZVByaW1hcnlFbnRpdHlJblJlbGF0aW9uc2hpcDogYW55O1xyXG4gICAgICAgIENhbkJlUmVsYXRlZEVudGl0eUluUmVsYXRpb25zaGlwOiBhbnk7XHJcbiAgICAgICAgQ2FuQ2hhbmdlSGllcmFyY2hpY2FsUmVsYXRpb25zaGlwOiBhbnk7XHJcbiAgICAgICAgQ2FuQ2hhbmdlVHJhY2tpbmdCZUVuYWJsZWQ6IGFueTtcclxuICAgICAgICBDYW5DcmVhdGVBdHRyaWJ1dGVzOiBhbnk7XHJcbiAgICAgICAgQ2FuQ3JlYXRlQ2hhcnRzOiBhbnk7XHJcbiAgICAgICAgQ2FuQ3JlYXRlRm9ybXM6IGFueTtcclxuICAgICAgICBDYW5DcmVhdGVWaWV3czogYW55O1xyXG4gICAgICAgIENhbkVuYWJsZVN5bmNUb0V4dGVybmFsU2VhcmNoSW5kZXg6IGFueTtcclxuICAgICAgICBDYW5Nb2RpZnlBZGRpdGlvbmFsU2V0dGluZ3M6IGFueTtcclxuICAgICAgICBDYW5UcmlnZ2VyV29ya2Zsb3c6IGFueTtcclxuICAgICAgICBDaGFuZ2VUcmFja2luZ0VuYWJsZWQ6IGFueTtcclxuICAgICAgICBDb2xsZWN0aW9uU2NoZW1hTmFtZTogYW55O1xyXG4gICAgICAgIERhdGFQcm92aWRlcklkOiBhbnk7XHJcbiAgICAgICAgRGF0YVNvdXJjZUlkOiBhbnk7XHJcbiAgICAgICAgRGF5c1NpbmNlUmVjb3JkTGFzdE1vZGlmaWVkOiBhbnk7XHJcbiAgICAgICAgRGVzY3JpcHRpb246IGFueTtcclxuICAgICAgICBEaXNwbGF5Q29sbGVjdGlvbk5hbWU6IGFueTtcclxuICAgICAgICBEaXNwbGF5TmFtZTogYW55O1xyXG4gICAgICAgIEVuZm9yY2VTdGF0ZVRyYW5zaXRpb25zOiBhbnk7XHJcbiAgICAgICAgRW50aXR5Q29sb3I6IGFueTtcclxuICAgICAgICBFbnRpdHlIZWxwVXJsOiBhbnk7XHJcbiAgICAgICAgRW50aXR5SGVscFVybEVuYWJsZWQ6IGFueTtcclxuICAgICAgICBFbnRpdHlTZXROYW1lOiBhbnk7XHJcbiAgICAgICAgRXh0ZXJuYWxDb2xsZWN0aW9uTmFtZTogYW55O1xyXG4gICAgICAgIEV4dGVybmFsTmFtZTogYW55O1xyXG4gICAgICAgIEhhc0FjdGl2aXRpZXM6IGFueTtcclxuICAgICAgICBIYXNDaGFuZ2VkOiBhbnk7XHJcbiAgICAgICAgSGFzRmVlZGJhY2s6IGFueTtcclxuICAgICAgICBIYXNOb3RlczogYW55O1xyXG4gICAgICAgIEljb25MYXJnZU5hbWU6IGFueTtcclxuICAgICAgICBJY29uTWVkaXVtTmFtZTogYW55O1xyXG4gICAgICAgIEljb25TbWFsbE5hbWU6IGFueTtcclxuICAgICAgICBJY29uVmVjdG9yTmFtZTogYW55O1xyXG4gICAgICAgIEludHJvZHVjZWRWZXJzaW9uOiBhbnk7XHJcbiAgICAgICAgSXNBSVJVcGRhdGVkOiBhbnk7XHJcbiAgICAgICAgSXNBY3Rpdml0eTogYW55O1xyXG4gICAgICAgIElzQWN0aXZpdHlQYXJ0eTogYW55O1xyXG4gICAgICAgIElzQXVkaXRFbmFibGVkOiBhbnk7XHJcbiAgICAgICAgSXNBdmFpbGFibGVPZmZsaW5lOiBhbnk7XHJcbiAgICAgICAgSXNCUEZFbnRpdHk6IGFueTtcclxuICAgICAgICBJc0J1c2luZXNzUHJvY2Vzc0VuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc0NoaWxkRW50aXR5OiBhbnk7XHJcbiAgICAgICAgSXNDb25uZWN0aW9uc0VuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc0N1c3RvbUVudGl0eTogYW55O1xyXG4gICAgICAgIElzQ3VzdG9taXphYmxlOiBhbnk7XHJcbiAgICAgICAgSXNEb2N1bWVudE1hbmFnZW1lbnRFbmFibGVkOiBhbnk7XHJcbiAgICAgICAgSXNEb2N1bWVudFJlY29tbWVuZGF0aW9uc0VuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc0R1cGxpY2F0ZURldGVjdGlvbkVuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc0VuYWJsZWRGb3JDaGFydHM6IGFueTtcclxuICAgICAgICBJc0VuYWJsZWRGb3JFeHRlcm5hbENoYW5uZWxzOiBhbnk7XHJcbiAgICAgICAgSXNFbmFibGVkRm9yVHJhY2U6IGFueTtcclxuICAgICAgICBJc0ltcG9ydGFibGU6IGFueTtcclxuICAgICAgICBJc0ludGVyYWN0aW9uQ2VudHJpY0VuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc0ludGVyc2VjdDogYW55O1xyXG4gICAgICAgIElzS25vd2xlZGdlTWFuYWdlbWVudEVuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc0xvZ2ljYWxFbnRpdHk6IGFueTtcclxuICAgICAgICBJc01TVGVhbXNJbnRlZ3JhdGlvbkVuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc01haWxNZXJnZUVuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc01hbmFnZWQ6IGFueTtcclxuICAgICAgICBJc01hcHBhYmxlOiBhbnk7XHJcbiAgICAgICAgSXNPZmZsaW5lSW5Nb2JpbGVDbGllbnQ6IGFueTtcclxuICAgICAgICBJc09uZU5vdGVJbnRlZ3JhdGlvbkVuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc09wdGltaXN0aWNDb25jdXJyZW5jeUVuYWJsZWQ6IGFueTtcclxuICAgICAgICBJc1ByaXZhdGU6IGFueTtcclxuICAgICAgICBJc1F1aWNrQ3JlYXRlRW5hYmxlZDogYW55O1xyXG4gICAgICAgIElzUmVhZE9ubHlJbk1vYmlsZUNsaWVudDogYW55O1xyXG4gICAgICAgIElzUmVhZGluZ1BhbmVFbmFibGVkOiBhbnk7XHJcbiAgICAgICAgSXNSZW5hbWVhYmxlOiBhbnk7XHJcbiAgICAgICAgSXNTTEFFbmFibGVkOiBhbnk7XHJcbiAgICAgICAgSXNTdGF0ZU1vZGVsQXdhcmU6IGFueTtcclxuICAgICAgICBJc1ZhbGlkRm9yQWR2YW5jZWRGaW5kOiBhbnk7XHJcbiAgICAgICAgSXNWYWxpZEZvclF1ZXVlOiBhbnk7XHJcbiAgICAgICAgSXNWaXNpYmxlSW5Nb2JpbGU6IGFueTtcclxuICAgICAgICBJc1Zpc2libGVJbk1vYmlsZUNsaWVudDogYW55O1xyXG4gICAgICAgIEtleXM6IGFueVtdO1xyXG4gICAgICAgIExvZ2ljYWxDb2xsZWN0aW9uTmFtZTogYW55O1xyXG4gICAgICAgIExvZ2ljYWxOYW1lOiBTdHJpbmc7XHJcbiAgICAgICAgTWFueVRvTWFueVJlbGF0aW9uc2hpcHM6IGFueVtdO1xyXG4gICAgICAgIE1hbnlUb09uZVJlbGF0aW9uc2hpcHM6IGFueVtdO1xyXG4gICAgICAgIE1ldGFkYXRhSWQ6IFN0cmluZztcclxuICAgICAgICBNb2JpbGVPZmZsaW5lRmlsdGVyczogYW55O1xyXG4gICAgICAgIE9iamVjdFR5cGVDb2RlOiBhbnk7XHJcbiAgICAgICAgT25lVG9NYW55UmVsYXRpb25zaGlwczogYW55W107XHJcbiAgICAgICAgT3duZXJzaGlwVHlwZTogYW55O1xyXG4gICAgICAgIFByaW1hcnlJZEF0dHJpYnV0ZTogYW55O1xyXG4gICAgICAgIFByaW1hcnlJbWFnZUF0dHJpYnV0ZTogYW55O1xyXG4gICAgICAgIFByaW1hcnlOYW1lQXR0cmlidXRlOiBhbnk7XHJcbiAgICAgICAgUHJpdmlsZWdlczogYW55W107XHJcbiAgICAgICAgUmVjdXJyZW5jZUJhc2VFbnRpdHlMb2dpY2FsTmFtZTogYW55O1xyXG4gICAgICAgIFJlcG9ydFZpZXdOYW1lOiBhbnk7XHJcbiAgICAgICAgU2NoZW1hTmFtZTogYW55O1xyXG4gICAgICAgIFN5bmNUb0V4dGVybmFsU2VhcmNoSW5kZXg6IGFueTtcclxuICAgICAgICBVc2VzQnVzaW5lc3NEYXRhTGFiZWxUYWJsZTogYW55O1xyXG4gICAgICB9W107XHJcbiAgICAgIFNlcnZlclZlcnNpb25TdGFtcDogU3RyaW5nO1xyXG4gICAgfSA9IGF3YWl0ICg8YW55PnJhd1Jlc3BvbnNlKS5qc29uKCk7XHJcblxyXG4gICAgYXNzZXJ0LmVxdWFsKFxyXG4gICAgICByZXNwb25zZS5FbnRpdHlNZXRhZGF0YVswXS5Mb2dpY2FsTmFtZSxcclxuICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgIFwiQWNjb3VudCBtZXRhZGF0YSByZXR1cm5lZFwiXHJcbiAgICApO1xyXG4gICAgYXNzZXJ0Lm9rKFxyXG4gICAgICByZXNwb25zZS5FbnRpdHlNZXRhZGF0YVswXS5BdHRyaWJ1dGVzLmxlbmd0aCA+IDAsXHJcbiAgICAgIFwiQWNjb3VudCBBdHRyaWJ1dGVzIHJldHVybmVkXCJcclxuICAgICk7XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWVzOlxyXG4vLyAgMS4gQ3JlYXRpbmcgYSByZWNvcmQgd2l0aCBhIGxvb2t1cCBmaWVsZFxyXG4vLyAgMi4gVXBkYXRpbmcgYSBsb29rdXAgZmllbGQgb24gYW4gZXhpc3RpbmcgcmVjb3JkXHJcbi8vICAzLiBTZXR0aW5nIGEgbG9va3VwIGZpZWxkIHRvIG51bGwgb24gYW4gZXhpc3RpbmcgcmVjb3JkXHJcbi8vICA0LiBRdWVyeWluZyBMb29rdXBzXHJcblxyXG4vLy8gPHJlZmVyZW5jZSBwYXRoPVwiLi4vV2ViQXBpUmVxdWVzdC50c1wiIC8+XHJcbmRlc2NyaWJlKFwiXCIsIGZ1bmN0aW9uKCkge1xyXG4gIGl0KFwiTG9va3VwIEZpZWxkc1wiLCBhc3luYyBmdW5jdGlvbigpIHtcclxuICAgIHRoaXMudGltZW91dCg5MDAwMCk7XHJcblxyXG4gICAgLy8gQ3JlYXRlIGEgY29udGFjdFxyXG4gICAgdmFyIGNvbnRhY3QxaWQ6IHtcclxuICAgICAgZW50aXR5VHlwZTogc3RyaW5nO1xyXG4gICAgICBpZDogc3RyaW5nO1xyXG4gICAgfSA9IGF3YWl0ICg8YW55PlhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiY29udGFjdFwiLCB7XHJcbiAgICAgIGxhc3RuYW1lOiBcIlNhbXBsZSBDb250YWN0IDFcIlxyXG4gICAgfSkpO1xyXG5cclxuICAgIHZhciBjb250YWN0MmlkOiB7XHJcbiAgICAgIGVudGl0eVR5cGU6IHN0cmluZztcclxuICAgICAgaWQ6IHN0cmluZztcclxuICAgIH0gPSBhd2FpdCAoPGFueT5Ycm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImNvbnRhY3RcIiwge1xyXG4gICAgICBsYXN0bmFtZTogXCJTYW1wbGUgQ29udGFjdCAyXCJcclxuICAgIH0pKTtcclxuXHJcbiAgICAvLyBDcmVhdGUgYWNjb3VudFxyXG4gICAgdmFyIGFjY291bnRpZDoge1xyXG4gICAgICBlbnRpdHlUeXBlOiBzdHJpbmc7XHJcbiAgICAgIGlkOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgKDxhbnk+WHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiLFxyXG4gICAgICBcInByaW1hcnljb250YWN0aWRAb2RhdGEuYmluZFwiOiBgL2NvbnRhY3RzKCR7Y29udGFjdDFpZC5pZH0pYFxyXG4gICAgfSkpO1xyXG5cclxuICAgIHRyeSB7XHJcbiAgICAgIC8vIFVwZGF0ZSB0aGUgcHJpbWFyeSBjb250YWN0IHRvIGEgZGlmZmVyZW50IGNvbnRhY3RcclxuICAgICAgYXdhaXQgWHJtLldlYkFwaS51cGRhdGVSZWNvcmQoXCJhY2NvdW50XCIsYWNjb3VudGlkLmlkLCB7XHJcbiAgICAgICAgXCJwcmltYXJ5Y29udGFjdGlkQG9kYXRhLmJpbmRcIjogYC9jb250YWN0cygke2NvbnRhY3QyaWQuaWR9KWBcclxuXHJcbiAgICAgIH0pO1xyXG5cclxuICAgICAgLy8gRGlzYXNzb2NpYXRlIENvbnRhY3QgdG8gQWNjb3VudCBhcyBQcmltYXJ5IENvbnRhY3RcclxuICAgICAgLy8gTm90ZTogIEl0IGlzIG5vdCBwb3NzaWJsZSB0byB1cGRhdGUgYSBsb29rdXAgZmllbGQgdG8gYmUgbnVsbFxyXG4gICAgICAvLyAgICAgICAgRWFjaCBmaWVsZCBiZWlnbiBudWxsZWQgbXVzdCBoYXZlIGEgc2VwYXJhdGUgREVMRVRFIHJlcXVlc3QgXHJcbiAgICAgIHZhciB1cmwgPSBgL2FjY291bnRzKCR7YWNjb3VudGlkLmlkfSkvcHJpbWFyeWNvbnRhY3RpZC8kcmVmYDtcclxuICAgICAgdmFyIHJlc3BvbnNlID0gYXdhaXQgV2ViQXBpUmVxdWVzdC5yZXF1ZXN0KFwiREVMRVRFXCIsdXJsKTtcclxuXHJcbiAgICB9IGZpbmFsbHkge1xyXG4gICAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgICBpZiAoYWNjb3VudGlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJhY2NvdW50XCIsIGFjY291bnRpZC5pZCk7XHJcbiAgICAgIH1cclxuXHJcbiAgICAgIC8vIERlbGV0ZSBjb250YWN0IDFcclxuICAgICAgaWYgKGNvbnRhY3QxaWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImNvbnRhY3RcIiwgY29udGFjdDFpZC5pZCk7XHJcbiAgICAgIH1cclxuXHJcbiAgICAgIC8vIERlbGV0ZSBjb250YWN0IDJcclxuICAgICAgaWYgKGNvbnRhY3QyaWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImNvbnRhY3RcIiwgY29udGFjdDJpZC5pZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZTpcclxuLy8gIDEuIEFzc29jaWF0aW5nIHR3byByZWNvcmRzIG92ZXIgYSBtYW55IHRvIG1hbnkgcmVsYXRpb25zaGlwXHJcbi8vICAyLiBEaXNhc3NvY2lhdGluZyB0aGUgdHdvIHJlY29yZHNcclxuLy8gIFNlZTogaHR0cHM6Ly9kb2NzLm1pY3Jvc29mdC5jb20vZW4tdXMvZHluYW1pY3MzNjUvY3VzdG9tZXItZW5nYWdlbWVudC9kZXZlbG9wZXIvd2ViYXBpL2Fzc29jaWF0ZS1kaXNhc3NvY2lhdGUtZW50aXRpZXMtdXNpbmctd2ViLWFwaVxyXG5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJNYW55IHRvIE1hbnlcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG5cclxuICAgIC8vIENyZWF0ZSBhY2NvdW50XHJcbiAgICB2YXIgYWNjb3VudGlkOiB7XHJcbiAgICAgIGVudGl0eVR5cGU6IHN0cmluZztcclxuICAgICAgaWQ6IHN0cmluZztcclxuICAgIH0gPSBhd2FpdCAoPGFueT5Ycm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImFjY291bnRcIiwge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBBY2NvdW50XCJcclxuICAgIH0pKTtcclxuXHJcbiAgICAvLyBDcmVhdGUgbGVhZFxyXG4gICAgdmFyIGxlYWRpZDoge1xyXG4gICAgICBlbnRpdHlUeXBlOiBzdHJpbmc7XHJcbiAgICAgIGlkOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgKDxhbnk+WHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJsZWFkXCIsIHtcclxuICAgICAgbGFzdG5hbWU6IFwiU2FtcGxlIExlYWRcIlxyXG4gICAgfSkpO1xyXG5cclxuICAgIHRyeSB7XHJcbiAgICAgIC8vIEFzc29jaWF0ZSBNYW55VG9NYW55XHJcbiAgICAgIHZhciBhc3NvY2lhdGUgPSB7XHJcbiAgICAgICAgXCJAb2RhdGEuY29udGV4dFwiOiBXZWJBcGlSZXF1ZXN0LmdldE9kYXRhQ29udGV4dCgpLFxyXG4gICAgICAgIFwiQG9kYXRhLmlkXCI6IGBsZWFkcygke2xlYWRpZC5pZH0pYFxyXG4gICAgICB9O1xyXG4gICAgICB2YXIgdXJsID0gYC9hY2NvdW50cygke2FjY291bnRpZC5pZH0pL2FjY291bnRsZWFkc19hc3NvY2lhdGlvbi8kcmVmYDtcclxuICAgICAgdmFyIHJlc3BvbnNlID0gYXdhaXQgV2ViQXBpUmVxdWVzdC5yZXF1ZXN0KFwiUE9TVFwiLCB1cmwsIGFzc29jaWF0ZSk7XHJcblxyXG4gICAgICAvLyBDaGVjayB0aGF0IHRoZSBhc3NvY2lhdGlvbiBoYXMgYmVlbiBtYWRlXHJcbiAgICAgIHZhciBmZXRjaCA9IGA8ZmV0Y2ggbm8tbG9jaz1cInRydWVcIiA+XHJcbiAgICAgIDxlbnRpdHkgbmFtZT1cImFjY291bnRcIiA+XHJcbiAgICAgICAgPGF0dHJpYnV0ZSBuYW1lPVwibmFtZVwiIC8+XHJcbiAgICAgICAgPGZpbHRlcj5cclxuICAgICAgICAgIDxjb25kaXRpb24gYXR0cmlidXRlPVwiYWNjb3VudGlkXCIgb3BlcmF0b3I9XCJlcVwiIHZhbHVlPVwiJHthY2NvdW50aWQuaWR9XCIgLz5cclxuICAgICAgICA8L2ZpbHRlcj5cclxuICAgICAgICA8bGluay1lbnRpdHkgbmFtZT1cImFjY291bnRsZWFkc1wiIGZyb209XCJhY2NvdW50aWRcIiB0bz1cImFjY291bnRpZFwiIGludGVyc2VjdD1cInRydWVcIiA+XHJcbiAgICAgICAgICA8YXR0cmlidXRlIG5hbWU9XCJuYW1lXCIgLz5cclxuICAgICAgICA8L2xpbmstZW50aXR5PlxyXG4gICAgICA8L2VudGl0eT5cclxuICAgIDwvZmV0Y2g+YDtcclxuXHJcbiAgICAgIHZhciBhc3NvY2lhdGVkTGVhZHMgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlTXVsdGlwbGVSZWNvcmRzKFxyXG4gICAgICAgIFwiYWNjb3VudFwiLFxyXG4gICAgICAgIFwiP2ZldGNoWG1sPVwiICsgZmV0Y2hcclxuICAgICAgKTtcclxuXHJcbiAgICAgIGFzc2VydC5lcXVhbChhc3NvY2lhdGVkTGVhZHMuZW50aXRpZXMubGVuZ3RoLCAxLCBcIkFzc29jaWF0ZWQgcmVjb3Jkc1wiKTtcclxuXHJcbiAgICAgIC8vIERpc2Fzc29jaWF0ZSBPbmVUb01hbnlcclxuICAgICAgdmFyIHVybCA9IGAvYWNjb3VudHMoJHthY2NvdW50aWQuaWR9KS9hY2NvdW50bGVhZHNfYXNzb2NpYXRpb24oJHtcclxuICAgICAgICBsZWFkaWQuaWRcclxuICAgICAgfSkvJHJlZmA7XHJcbiAgICAgIHZhciByZXNwb25zZSA9IGF3YWl0IFdlYkFwaVJlcXVlc3QucmVxdWVzdChcIkRFTEVURVwiLCB1cmwpO1xyXG4gICAgfSBmaW5hbGx5IHtcclxuICAgICAgaWYgKGxlYWRpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwibGVhZFwiLCBsZWFkaWQuaWQpO1xyXG4gICAgICB9XHJcbiAgICAgIC8vIERlbGV0ZSBhY2NvdW50XHJcbiAgICAgIGlmIChhY2NvdW50aWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudGlkLmlkKTtcclxuICAgICAgfVxyXG4gICAgfVxyXG4gIH0pO1xyXG59KTtcclxuIiwiLy8gRGVtb25zdHJhdGVzIHRoZSBmb2xsb3dpbmcgdGVjaG5pcXVlOlxyXG4vLyAgMS4gQXNzb2NpYXRpbmcgdHdvIHJlY29yZHMgb3ZlciBhIG1hbnktdG8tbWFueSByZWxhdGlvbnNoaXBcclxuLy8gIDIuIERpc2Fzc29jaWF0aW5nIHRoZSB0d28gcmVjb3Jkc1xyXG4vLyAgTk9URTogVGhpcyBzYW1wbGUgdXNlcyB0aGUgZXhlY3V0ZSBtZXRob2Qgd2l0aCBhIG9wZXJhdGlvbk5hbWUgb2YgJ0Fzc29jaWF0ZScgYW5kICdEaXNhc3NvY2lhdGUnXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIk1hbnkgdG8gT25lIHVzaW5nIGV4ZWN1dGVcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgLy8gQ3JlYXRlIGEgY29udGFjdFxyXG4gICAgdmFyIGNvbnRhY3RpZDoge1xyXG4gICAgICBlbnRpdHlUeXBlOiBzdHJpbmc7XHJcbiAgICAgIGlkOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgKDxhbnk+WHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJjb250YWN0XCIsIHtcclxuICAgICAgbGFzdG5hbWU6IFwiU2FtcGxlIENvbnRhY3RcIlxyXG4gICAgfSkpO1xyXG5cclxuICAgIC8vIENyZWF0ZSBhY2NvdW50XHJcbiAgICB2YXIgYWNjb3VudGlkOiB7XHJcbiAgICAgIGVudGl0eVR5cGU6IHN0cmluZztcclxuICAgICAgaWQ6IHN0cmluZztcclxuICAgIH0gPSBhd2FpdCAoPGFueT5Ycm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImFjY291bnRcIiwge1xyXG4gICAgICBuYW1lOiBcIlNhbXBsZSBBY2NvdW50XCJcclxuICAgIH0pKTtcclxuXHJcbiAgICB0cnkge1xyXG4gICAgICAvLyBBc3NvY2lhdGUgQ29udGFjdCB0byBBY2NvdW50IGFzIFByaW1hcnkgQ29udGFjdFxyXG4gICAgICB2YXIgYXNzb2NpYXRlUmVxdWVzdCA9IG5ldyBjbGFzcyB7XHJcbiAgICAgICAgdGFyZ2V0ID0ge1xyXG4gICAgICAgICAgaWQ6IGFjY291bnRpZC5pZCxcclxuICAgICAgICAgIGVudGl0eVR5cGU6IFwiYWNjb3VudFwiXHJcbiAgICAgICAgfTtcclxuICAgICAgICByZWxhdGVkRW50aXRpZXMgPSBbXHJcbiAgICAgICAgICB7XHJcbiAgICAgICAgICAgIGlkOiBjb250YWN0aWQuaWQsXHJcbiAgICAgICAgICAgIGVudGl0eVR5cGU6IFwiY29udGFjdFwiXHJcbiAgICAgICAgICB9XHJcbiAgICAgICAgXTtcclxuICAgICAgICByZWxhdGlvbnNoaXAgPSBcImFjY291bnRfcHJpbWFyeV9jb250YWN0XCI7XHJcbiAgICAgICAgZ2V0TWV0YWRhdGEoKTogYW55IHtcclxuICAgICAgICAgIHJldHVybiB7XHJcbiAgICAgICAgICAgIHBhcmFtZXRlclR5cGVzOiB7fSxcclxuICAgICAgICAgICAgb3BlcmF0aW9uVHlwZTogMixcclxuICAgICAgICAgICAgb3BlcmF0aW9uTmFtZTogXCJBc3NvY2lhdGVcIlxyXG4gICAgICAgICAgfTtcclxuICAgICAgICB9XHJcbiAgICAgIH0oKTtcclxuXHJcbiAgICAgIHZhciByZXNwb25zZSA9IGF3YWl0IFhybS5XZWJBcGkub25saW5lLmV4ZWN1dGUoYXNzb2NpYXRlUmVxdWVzdCk7XHJcblxyXG4gICAgICAvLyBEaXNhc3NvY2lhdGUgQ29udGFjdCB0byBBY2NvdW50IGFzIFByaW1hcnkgQ29udGFjdFxyXG4gICAgICAvLyBOb3RlOiBXaGVuIGRpc2Fzc29jaWF0aW5nIE1hbnkgdG8gT25lIC0gdGhlIGxvb2t1cCBhdHRyaWJ1dGUgbmFtZSBpcyB1c2VkIHdpdGggb25seSB0aGUgbGVmdCBoYW5kIHNpZGUgaWRcclxuICAgICAgdmFyIGRpc3Nhc3NvY2lhdGVSZXF1ZXN0ID0gbmV3IGNsYXNzIHtcclxuICAgICAgICB0YXJnZXQgPSB7XHJcbiAgICAgICAgICBpZDogYWNjb3VudGlkLmlkLFxyXG4gICAgICAgICAgZW50aXR5VHlwZTogXCJhY2NvdW50XCJcclxuICAgICAgICB9O1xyXG4gICAgICAgIHJlbGF0aW9uc2hpcCA9IFwicHJpbWFyeWNvbnRhY3RpZFwiO1xyXG4gICAgICAgIGdldE1ldGFkYXRhKCk6IGFueSB7XHJcbiAgICAgICAgICByZXR1cm4ge1xyXG4gICAgICAgICAgICBwYXJhbWV0ZXJUeXBlczoge30sXHJcbiAgICAgICAgICAgIG9wZXJhdGlvblR5cGU6IDIsXHJcbiAgICAgICAgICAgIG9wZXJhdGlvbk5hbWU6IFwiRGlzYXNzb2NpYXRlXCJcclxuICAgICAgICAgIH07XHJcbiAgICAgICAgfVxyXG4gICAgICB9KCk7XHJcblxyXG4gICAgICB2YXIgZGlzc2Fzc29jaWF0ZVJlc3BvbnNlID0gYXdhaXQgWHJtLldlYkFwaS5vbmxpbmUuZXhlY3V0ZShcclxuICAgICAgICBkaXNzYXNzb2NpYXRlUmVxdWVzdFxyXG4gICAgICApO1xyXG5cclxuICAgIH0gZmluYWxseSB7XHJcblxyXG4gICAgICAvLyBEZWxldGUgYWNjb3VudFxyXG4gICAgICBpZiAoYWNjb3VudGlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJhY2NvdW50XCIsIGFjY291bnRpZC5pZCk7XHJcbiAgICAgIH1cclxuXHJcbiAgICAgIC8vIERlbGV0ZSBjb250YWN0XHJcbiAgICAgIGlmIChjb250YWN0aWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImNvbnRhY3RcIiwgY29udGFjdGlkLmlkKTtcclxuICAgICAgfVxyXG4gICAgfVxyXG4gIH0pO1xyXG59KTtcclxuIiwiLy8vIDxyZWZlcmVuY2UgcGF0aD1cIi4uL1dlYkFwaVJlcXVlc3QudHNcIiAvPlxyXG4vLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWU6XHJcbi8vICAxLiBBc3NvY2lhdGluZyB0d28gcmVjb3JkcyBvdmVyIGEgbWFueS10by1vbmUgcmVsYXRpb25zaGlwLiBUaGlzIGlzIGFuIGFsdGVybmF0aXZlIHRvIHNpbXBseSB1c2luZyBhIGxvb2t1cCBmaWVsZCBpbiBhIGNyZWF0ZS91cGRhdGVcclxuLy8gIDIuIERpc2Fzc29jaWF0aW5nIHRoZSB0d28gcmVjb3Jkc1xyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvYXNzb2NpYXRlLWRpc2Fzc29jaWF0ZS1lbnRpdGllcy11c2luZy13ZWItYXBpXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIk1hbnkgdG8gT25lXCIsIGFzeW5jIGZ1bmN0aW9uKCkge1xyXG4gICAgdGhpcy50aW1lb3V0KDkwMDAwKTtcclxuICAgIC8vIENyZWF0ZSBhIGNvbnRhY3RcclxuICAgIHZhciBjb250YWN0aWQ6IHtcclxuICAgICAgZW50aXR5VHlwZTogc3RyaW5nO1xyXG4gICAgICBpZDogc3RyaW5nO1xyXG4gICAgfSA9IGF3YWl0ICg8YW55PlhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiY29udGFjdFwiLCB7XHJcbiAgICAgIGxhc3RuYW1lOiBcIlNhbXBsZSBDb250YWN0XCJcclxuICAgIH0pKTtcclxuXHJcbiAgICAvLyBDcmVhdGUgYWNjb3VudFxyXG4gICAgdmFyIGFjY291bnRpZDoge1xyXG4gICAgICBlbnRpdHlUeXBlOiBzdHJpbmc7XHJcbiAgICAgIGlkOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgKDxhbnk+WHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiXHJcbiAgICB9KSk7XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgLy8gQXNzb2NpYXRlIENvbnRhY3QgdG8gQWNjb3VudCBhcyBQcmltYXJ5IENvbnRhY3RcclxuICAgICAgdmFyIGFzc29jaWF0ZSA9IHtcclxuICAgICAgICBcIkBvZGF0YS5jb250ZXh0XCI6IFdlYkFwaVJlcXVlc3QuZ2V0T2RhdGFDb250ZXh0KCksXHJcbiAgICAgICAgXCJAb2RhdGEuaWRcIjogYGNvbnRhY3RzKCR7Y29udGFjdGlkLmlkfSlgXHJcbiAgICAgIH07XHJcbiAgICAgIHZhciB1cmwgPSBgL2FjY291bnRzKCR7YWNjb3VudGlkLmlkfSkvcHJpbWFyeWNvbnRhY3RpZC8kcmVmYDtcclxuICAgICAgdmFyIHJlc3BvbnNlID0gYXdhaXQgV2ViQXBpUmVxdWVzdC5yZXF1ZXN0KFwiUFVUXCIsIHVybCwgYXNzb2NpYXRlKTtcclxuXHJcbiAgICAgIC8vIERpc2Fzc29jaWF0ZSBDb250YWN0IHRvIEFjY291bnQgYXMgUHJpbWFyeSBDb250YWN0XHJcbiAgICAgIC8vIE5vdGU6IFdoZW4gZGlzYXNzb2NpYXRpbmcgTWFueSB0byBPbmUgLSB0aGUgbG9va3VwIGF0dHJpYnV0ZSBuYW1lIGlzIHVzZWQgd2l0aCBvbmx5IHRoZSBsZWZ0IGhhbmQgc2lkZSBpZFxyXG4gICAgICB2YXIgdXJsID0gYC9hY2NvdW50cygke2FjY291bnRpZC5pZH0pL3ByaW1hcnljb250YWN0aWQvJHJlZmA7XHJcbiAgICAgIHZhciByZXNwb25zZSA9IGF3YWl0IFdlYkFwaVJlcXVlc3QucmVxdWVzdChcIkRFTEVURVwiLHVybCk7XHJcblxyXG4gICAgfSBmaW5hbGx5IHtcclxuICAgICAgLy8gRGVsZXRlIGFjY291bnRcclxuICAgICAgaWYgKGFjY291bnRpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiYWNjb3VudFwiLCBhY2NvdW50aWQuaWQpO1xyXG4gICAgICB9XHJcblxyXG4gICAgICAvLyBEZWxldGUgY29udGFjdFxyXG4gICAgICBpZiAoY29udGFjdGlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJjb250YWN0XCIsIGNvbnRhY3RpZC5pZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZTpcclxuLy8gIDEuIEFzc29jaWF0aW5nIHR3byByZWNvcmRzIG92ZXIgYSBvbmUtdG8tbWFueSByZWxhdGlvbnNoaXBcclxuLy8gIDIuIERpc2Fzc29jaWF0aW5nIHRoZSB0d28gcmVjb3Jkc1xyXG4vLyAgTk9URTogVGhpcyBzYW1wbGUgdXNlcyB0aGUgZXhlY3V0ZSBtZXRob2Qgd2l0aCBhIG9wZXJhdGlvbk5hbWUgb2YgJ0Fzc29jaWF0ZScgYW5kICdEaXNhc3NvY2lhdGUnXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIk9uZSB0byBNYW55IHVzaW5nIGV4ZWN1dGVcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG4gICAgXHJcbiAgICAvLyBDcmVhdGUgYSBjb250YWN0XHJcbiAgICB2YXIgY29udGFjdGlkOiB7XHJcbiAgICAgIGVudGl0eVR5cGU6IHN0cmluZztcclxuICAgICAgaWQ6IHN0cmluZztcclxuICAgIH0gPSBhd2FpdCAoPGFueT5Ycm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImNvbnRhY3RcIiwge1xyXG4gICAgICBsYXN0bmFtZTogXCJTYW1wbGUgQ29udGFjdFwiXHJcbiAgICB9KSk7XHJcblxyXG4gICAgLy8gQ3JlYXRlIGFjY291bnRcclxuICAgIHZhciBhY2NvdW50aWQ6IHtcclxuICAgICAgZW50aXR5VHlwZTogc3RyaW5nO1xyXG4gICAgICBpZDogc3RyaW5nO1xyXG4gICAgfSA9IGF3YWl0ICg8YW55PlhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiYWNjb3VudFwiLCB7XHJcbiAgICAgIG5hbWU6IFwiU2FtcGxlIEFjY291bnRcIlxyXG4gICAgfSkpO1xyXG5cclxuICAgIHRyeSB7XHJcbiAgICAgIC8vIEFzc29jaWF0ZSBDb250YWN0IHRvIEFjY291bnQgYXMgUHJpbWFyeSBDb250YWN0XHJcbiAgICAgIHZhciBhc3NvY2lhdGVSZXF1ZXN0ID0gbmV3IGNsYXNzIHtcclxuICAgICAgICB0YXJnZXQgPSB7XHJcbiAgICAgICAgICBpZDogY29udGFjdGlkLmlkLFxyXG4gICAgICAgICAgZW50aXR5VHlwZTogXCJjb250YWN0XCJcclxuICAgICAgICB9O1xyXG4gICAgICAgIHJlbGF0ZWRFbnRpdGllcyA9IFtcclxuICAgICAgICAgIHtcclxuICAgICAgICAgICAgaWQ6IGFjY291bnRpZC5pZCxcclxuICAgICAgICAgICAgZW50aXR5VHlwZTogXCJhY2NvdW50XCJcclxuICAgICAgICAgIH1cclxuICAgICAgICBdO1xyXG4gICAgICAgIHJlbGF0aW9uc2hpcCA9IFwiYWNjb3VudF9wcmltYXJ5X2NvbnRhY3RcIjtcclxuICAgICAgICBnZXRNZXRhZGF0YSgpOiBhbnkge1xyXG4gICAgICAgICAgcmV0dXJuIHtcclxuICAgICAgICAgICAgcGFyYW1ldGVyVHlwZXM6IHt9LFxyXG4gICAgICAgICAgICBvcGVyYXRpb25UeXBlOiAyLFxyXG4gICAgICAgICAgICBvcGVyYXRpb25OYW1lOiBcIkFzc29jaWF0ZVwiXHJcbiAgICAgICAgICB9O1xyXG4gICAgICAgIH1cclxuICAgICAgfSgpO1xyXG5cclxuICAgICAgdmFyIHJlc3BvbnNlID0gYXdhaXQgWHJtLldlYkFwaS5vbmxpbmUuZXhlY3V0ZShhc3NvY2lhdGVSZXF1ZXN0KTtcclxuXHJcbiAgICAgIC8vIERpc2Fzc29jaWF0ZSBDb250YWN0IHRvIEFjY291bnQgYXMgUHJpbWFyeSBDb250YWN0XHJcbiAgICAgIHZhciBkaXNzYXNzb2NpYXRlUmVxdWVzdCA9IG5ldyBjbGFzcyB7XHJcbiAgICAgICAgdGFyZ2V0ID0ge1xyXG4gICAgICAgICAgaWQ6IGNvbnRhY3RpZC5pZCxcclxuICAgICAgICAgIGVudGl0eVR5cGU6IFwiY29udGFjdFwiXHJcbiAgICAgICAgfTtcclxuICAgICAgICByZWxhdGVkRW50aXR5SWQgPSBhY2NvdW50aWQuaWQ7XHJcbiAgICAgICAgcmVsYXRpb25zaGlwID0gXCJhY2NvdW50X3ByaW1hcnlfY29udGFjdFwiO1xyXG4gICAgICAgIGdldE1ldGFkYXRhKCk6IGFueSB7XHJcbiAgICAgICAgICByZXR1cm4ge1xyXG4gICAgICAgICAgICBwYXJhbWV0ZXJUeXBlczoge30sXHJcbiAgICAgICAgICAgIG9wZXJhdGlvblR5cGU6IDIsXHJcbiAgICAgICAgICAgIG9wZXJhdGlvbk5hbWU6IFwiRGlzYXNzb2NpYXRlXCJcclxuICAgICAgICAgIH07XHJcbiAgICAgICAgfVxyXG4gICAgICB9KCk7XHJcblxyXG4gICAgICB2YXIgZGlzc2Fzc29jaWF0ZVJlc3BvbnNlID0gYXdhaXQgWHJtLldlYkFwaS5vbmxpbmUuZXhlY3V0ZShcclxuICAgICAgICBkaXNzYXNzb2NpYXRlUmVxdWVzdFxyXG4gICAgICApO1xyXG4gICAgICBcclxuICAgIH0gZmluYWxseSB7XHJcbiAgICAgIC8vIERlbGV0ZSBhY2NvdW50XHJcbiAgICAgIGlmIChhY2NvdW50aWQpIHtcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudGlkLmlkKTtcclxuICAgICAgfVxyXG5cclxuICAgICAgLy8gRGVsZXRlIGNvbnRhY3RcclxuICAgICAgaWYgKGNvbnRhY3RpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiY29udGFjdFwiLCBjb250YWN0aWQuaWQpO1xyXG4gICAgICB9XHJcbiAgICB9XHJcbiAgfSk7XHJcbn0pO1xyXG4iLCIvLy8gPHJlZmVyZW5jZSBwYXRoPVwiLi4vV2ViQXBpUmVxdWVzdC50c1wiIC8+XHJcbi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZTpcclxuLy8gIDEuIEFzc29jaWF0aW5nIHR3byByZWNvcmRzIG92ZXIgYSBvbmUtdG8tbWFueSByZWxhdGlvbnNoaXBcclxuLy8gIDIuIERpc2Fzc29jaWF0aW5nIHRoZSB0d28gcmVjb3Jkc1xyXG4vLyAgU2VlOiBodHRwczovL2RvY3MubWljcm9zb2Z0LmNvbS9lbi11cy9keW5hbWljczM2NS9jdXN0b21lci1lbmdhZ2VtZW50L2RldmVsb3Blci93ZWJhcGkvYXNzb2NpYXRlLWRpc2Fzc29jaWF0ZS1lbnRpdGllcy11c2luZy13ZWItYXBpXHJcblxyXG5kZXNjcmliZShcIlwiLCBmdW5jdGlvbigpIHtcclxuICBpdChcIk9uZSB0byBNYW55XCIsIGFzeW5jIGZ1bmN0aW9uKCkge1xyXG4gICAgdGhpcy50aW1lb3V0KDkwMDAwKTtcclxuICAgIC8vIENyZWF0ZSBhIGNvbnRhY3RcclxuICAgIHZhciBjb250YWN0aWQ6IHtcclxuICAgICAgZW50aXR5VHlwZTogc3RyaW5nO1xyXG4gICAgICBpZDogc3RyaW5nO1xyXG4gICAgfSA9IGF3YWl0ICg8YW55PlhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiY29udGFjdFwiLCB7XHJcbiAgICAgIGxhc3RuYW1lOiBcIlNhbXBsZSBDb250YWN0XCJcclxuICAgIH0pKTtcclxuXHJcbiAgICAvLyBDcmVhdGUgYWNjb3VudFxyXG4gICAgdmFyIGFjY291bnRpZDoge1xyXG4gICAgICBlbnRpdHlUeXBlOiBzdHJpbmc7XHJcbiAgICAgIGlkOiBzdHJpbmc7XHJcbiAgICB9ID0gYXdhaXQgKDxhbnk+WHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJhY2NvdW50XCIsIHtcclxuICAgICAgbmFtZTogXCJTYW1wbGUgQWNjb3VudFwiXHJcbiAgICB9KSk7XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgLy8gQXNzb2NpYXRlIENvbnRhY3QgdG8gQWNjb3VudCBhcyBQcmltYXJ5IENvbnRhY3RcclxuICAgICAgdmFyIGFzc29jaWF0ZSA9IHtcclxuICAgICAgICBcIkBvZGF0YS5jb250ZXh0XCI6IFdlYkFwaVJlcXVlc3QuZ2V0T2RhdGFDb250ZXh0KCksXHJcbiAgICAgICAgXCJAb2RhdGEuaWRcIjogYGFjY291bnRzKCR7YWNjb3VudGlkLmlkfSlgXHJcbiAgICAgIH07XHJcbiAgICAgIHZhciB1cmwgPSBgL2NvbnRhY3RzKCR7Y29udGFjdGlkLmlkfSkvYWNjb3VudF9wcmltYXJ5X2NvbnRhY3QvJHJlZmA7XHJcbiAgICAgIHZhciByZXNwb25zZSA9IGF3YWl0IFdlYkFwaVJlcXVlc3QucmVxdWVzdChcIlBPU1RcIiwgdXJsLCBhc3NvY2lhdGUpO1xyXG5cclxuICAgICAgLy8gRGlzYXNzb2NpYXRlIENvbnRhY3QgdG8gQWNjb3VudCBhcyBQcmltYXJ5IENvbnRhY3RcclxuICAgICAgdmFyIHVybCA9IGAvY29udGFjdHMoJHtjb250YWN0aWQuaWR9KS9hY2NvdW50X3ByaW1hcnlfY29udGFjdCgke2FjY291bnRpZC5pZH0pLyRyZWZgO1xyXG4gICAgICB2YXIgcmVzcG9uc2UgPSBhd2FpdCBXZWJBcGlSZXF1ZXN0LnJlcXVlc3QoXCJERUxFVEVcIix1cmwpO1xyXG4gICAgIFxyXG4gICAgfSBmaW5hbGx5IHtcclxuICAgICAgLy8gRGVsZXRlIGFjY291bnRcclxuICAgICAgaWYgKGFjY291bnRpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiYWNjb3VudFwiLCBhY2NvdW50aWQuaWQpO1xyXG4gICAgICB9XHJcblxyXG4gICAgICAvLyBEZWxldGUgY29udGFjdFxyXG4gICAgICBpZiAoY29udGFjdGlkKSB7XHJcbiAgICAgICAgYXdhaXQgWHJtLldlYkFwaS5kZWxldGVSZWNvcmQoXCJjb250YWN0XCIsIGNvbnRhY3RpZC5pZCk7XHJcbiAgICAgIH1cclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiIsIi8vIERlbW9uc3RyYXRlcyB0aGUgZm9sbG93aW5nIHRlY2huaXF1ZXM6XHJcbi8vICAxLiBDcmVhdGluZyBhY3Rpdml0aWVzIHdpdGggYWN0aXZpdHkgcGFydGllc1xyXG4vLyAgMi4gVXBkYXRpbmcgYWN0aXZpdHkgcGFydGllcyAtIHRoaXMgaXMgYSBzcGVjaWFsIGNhc2VcclxuLy8gIFNlZTogaHR0cHM6Ly9kb2NzLm1pY3Jvc29mdC5jb20vZW4tdXMvZHluYW1pY3MzNjUvY3VzdG9tZXItZW5nYWdlbWVudC9kZXZlbG9wZXIvd2ViYXBpL2Fzc29jaWF0ZS1kaXNhc3NvY2lhdGUtZW50aXRpZXMtdXNpbmctd2ViLWFwaSNhc3NvY2lhdGUtZW50aXRpZXMtb24tdXBkYXRlLXVzaW5nLWNvbGxlY3Rpb24tdmFsdWVkLW5hdmlnYXRpb24tcHJvcGVydHlcclxuXHJcbi8vLyA8cmVmZXJlbmNlIHBhdGg9XCIuLi9XZWJBcGlSZXF1ZXN0LnRzXCIgLz5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJBY3Rpdml0eSBQYXJ0aWVzIExldHRlclwiLCBhc3luYyBmdW5jdGlvbigpIHtcclxuICAgIHRoaXMudGltZW91dCg5MDAwMDApO1xyXG4gICAgdmFyIGFzc2VydCA9IGNoYWkuYXNzZXJ0O1xyXG5cclxuICAgIC8vIENyZWF0ZSBjb250YWN0IDFcclxuICAgIHZhciBjb250YWN0MSA9IHtcclxuICAgICAgbGFzdG5hbWU6IGBUZXN0IENvbnRhY3QgMSR7bmV3IERhdGUoKS50b1VUQ1N0cmluZygpfWBcclxuICAgIH07XHJcbiAgICB2YXIgY29udGFjdDFpZCA9ICg8YW55PmF3YWl0IFhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiY29udGFjdFwiLCBjb250YWN0MSkpXHJcbiAgICAgIC5pZDtcclxuXHJcbiAgICAvLyBDcmVhdGUgY29udGFjdCAyXHJcbiAgICB2YXIgY29udGFjdDIgPSB7XHJcbiAgICAgIGxhc3RuYW1lOiBgVGVzdCBDb250YWN0IDIgJHtuZXcgRGF0ZSgpLnRvVVRDU3RyaW5nKCl9YFxyXG4gICAgfTtcclxuICAgIHZhciBjb250YWN0MmlkID0gKDxhbnk+YXdhaXQgWHJtLldlYkFwaS5jcmVhdGVSZWNvcmQoXCJjb250YWN0XCIsIGNvbnRhY3QyKSlcclxuICAgICAgLmlkO1xyXG5cclxuICAgIHRyeSB7XHJcbiAgICAgIC8vIENyZWF0ZSBsZXR0ZXJcclxuICAgICAgY29uc3QgbGV0dGVyMToge1xyXG4gICAgICAgIHN1YmplY3Q6IHN0cmluZztcclxuICAgICAgICBsZXR0ZXJfYWN0aXZpdHlfcGFydGllczoge1xyXG4gICAgICAgICAgcGFydGljaXBhdGlvbnR5cGVtYXNrOiBudW1iZXI7XHJcbiAgICAgICAgICBbc29tZXRoaW5nOiBzdHJpbmddOiBhbnk7XHJcbiAgICAgICAgfVtdO1xyXG4gICAgICB9ID0ge1xyXG4gICAgICAgIHN1YmplY3Q6IGBTYW1wbGUgTGV0dGVyICR7bmV3IERhdGUoKS50b1VUQ1N0cmluZygpfWAsXHJcbiAgICAgICAgbGV0dGVyX2FjdGl2aXR5X3BhcnRpZXM6IFtcclxuICAgICAgICAgIHtcclxuICAgICAgICAgICAgcGFydGljaXBhdGlvbnR5cGVtYXNrOiAyLCAvLyBUb1xyXG4gICAgICAgICAgICBcIkBvZGF0YS50eXBlXCI6IFwiTWljcm9zb2Z0LkR5bmFtaWNzLkNSTS5hY3Rpdml0eXBhcnR5XCIsXHJcbiAgICAgICAgICAgIFwicGFydHlpZF9jb250YWN0QG9kYXRhLmJpbmRcIjogYGNvbnRhY3RzKCR7Y29udGFjdDFpZH0pYFxyXG4gICAgICAgICAgfVxyXG4gICAgICAgIF1cclxuICAgICAgfTtcclxuXHJcbiAgICAgIHZhciBsZXR0ZXIxaWQgPSAoPGFueT5hd2FpdCBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImxldHRlclwiLCBsZXR0ZXIxKSlcclxuICAgICAgICAuaWQ7XHJcblxyXG4gICAgICBpZiAoIWxldHRlcjFpZCkgdGhyb3cgbmV3IEVycm9yKFwiTGV0dGVyIG5vdCBjcmVhdGVkXCIpO1xyXG5cclxuICAgICAgLy8gUXVlcnkgdGhlIGxldHRlciBhbmQgY2hlY2sgdGhlIGF0dHJpYnV0ZSB2YWx1ZXNcclxuICAgICAgdmFyIGxldHRlcjIgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFxyXG4gICAgICAgIFwibGV0dGVyXCIsXHJcbiAgICAgICAgbGV0dGVyMWlkLFxyXG4gICAgICAgIFwiPyRleHBhbmQ9bGV0dGVyX2FjdGl2aXR5X3BhcnRpZXMoJHNlbGVjdD1hY3Rpdml0eXBhcnR5aWQsX3BhcnR5aWRfdmFsdWUscGFydGljaXBhdGlvbnR5cGVtYXNrKVwiXHJcbiAgICAgICk7XHJcblxyXG4gICAgICBpZiAoXHJcbiAgICAgICAgIWxldHRlcjIubGV0dGVyX2FjdGl2aXR5X3BhcnRpZXMgfHxcclxuICAgICAgICAhbGV0dGVyMS5sZXR0ZXJfYWN0aXZpdHlfcGFydGllcy5sZW5ndGhcclxuICAgICAgKVxyXG4gICAgICAgIHRocm93IG5ldyBFcnJvcihcIkxldHRlcjEgbGV0dGVyX2FjdGl2aXR5X3BhcnRpZXMgbm90IHNldFwiKTtcclxuXHJcbiAgICAgIHZhciBwYXJ0eVRvOiBhbnkgPSBmaW5kUGFydHlCeUlkKFxyXG4gICAgICAgIGxldHRlcjIubGV0dGVyX2FjdGl2aXR5X3BhcnRpZXMsXHJcbiAgICAgICAgY29udGFjdDFpZFxyXG4gICAgICApO1xyXG5cclxuICAgICAgYXNzZXJ0LmlzTm90TnVsbChwYXJ0eVRvLCBcIlRvIFBhcnR5IHJldHVybmVkXCIpO1xyXG5cclxuICAgICAgLy8gQWRkIGFuIGFjdGl2aXR5IHBhcnR5XHJcbiAgICAgIGxldHRlcjEubGV0dGVyX2FjdGl2aXR5X3BhcnRpZXMucHVzaCh7XHJcbiAgICAgICAgcGFydGljaXBhdGlvbnR5cGVtYXNrOiAyLCAvLyBUb1xyXG4gICAgICAgIFwiQG9kYXRhLnR5cGVcIjogXCJNaWNyb3NvZnQuRHluYW1pY3MuQ1JNLmFjdGl2aXR5cGFydHlcIixcclxuICAgICAgICBcInBhcnR5aWRfY29udGFjdEBvZGF0YS5iaW5kXCI6IGBjb250YWN0cygke2NvbnRhY3QyaWR9KWBcclxuICAgICAgfSk7XHJcblxyXG4gICAgICAvLyBVcGRhdGUgbGV0dGVyXHJcbiAgICAgIGF3YWl0IFhybS5XZWJBcGkudXBkYXRlUmVjb3JkKFwibGV0dGVyXCIsIGxldHRlcjFpZCwgbGV0dGVyMSk7XHJcblxyXG4gICAgICAvLyBRdWVyeSB0aGUgbGV0dGVyIGFuZCBjaGVjayB0aGUgYXR0cmlidXRlIHZhbHVlc1xyXG4gICAgICB2YXIgbGV0dGVyMyA9IGF3YWl0IFhybS5XZWJBcGkucmV0cmlldmVSZWNvcmQoXHJcbiAgICAgICAgXCJsZXR0ZXJcIixcclxuICAgICAgICBsZXR0ZXIxaWQsXHJcbiAgICAgICAgXCI/JGV4cGFuZD1sZXR0ZXJfYWN0aXZpdHlfcGFydGllcygkc2VsZWN0PWFjdGl2aXR5cGFydHlpZCxfcGFydHlpZF92YWx1ZSxwYXJ0aWNpcGF0aW9udHlwZW1hc2spXCJcclxuICAgICAgKTtcclxuXHJcbiAgICAgIHZhciBwYXJ0eVRvMjogYW55ID0gZmluZFBhcnR5QnlJZChcclxuICAgICAgICBsZXR0ZXIzLmxldHRlcl9hY3Rpdml0eV9wYXJ0aWVzLFxyXG4gICAgICAgIGNvbnRhY3QxaWRcclxuICAgICAgKTtcclxuICAgICAgdmFyIHBhcnR5VG8zOiBhbnkgPSBmaW5kUGFydHlCeUlkKFxyXG4gICAgICAgIGxldHRlcjMubGV0dGVyX2FjdGl2aXR5X3BhcnRpZXMsXHJcbiAgICAgICAgY29udGFjdDJpZFxyXG4gICAgICApO1xyXG5cclxuICAgICAgYXNzZXJ0LmlzTm90TnVsbChwYXJ0eVRvMiwgXCJUbyBQYXJ0eSAxIHJldHVybmVkXCIpO1xyXG4gICAgICBhc3NlcnQuaXNOb3ROdWxsKHBhcnR5VG8zLCBcIlRvIFBhcnR5IDIgcmV0dXJuZWRcIik7XHJcbiAgICB9IGZpbmFsbHkge1xyXG4gICAgICAvLyBEZWxldGUgQ29udGFjdFxyXG4gICAgICBpZiAoY29udGFjdDFpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiY29udGFjdFwiLCBjb250YWN0MWlkKTtcclxuICAgICAgfVxyXG4gICAgICBpZiAoY29udGFjdDJpZCkge1xyXG4gICAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiY29udGFjdFwiLCBjb250YWN0MmlkKTtcclxuICAgICAgfVxyXG4gICAgfVxyXG4gIH0pO1xyXG5cclxuICBmdW5jdGlvbiBmaW5kUGFydHlCeUlkKHBhcnRpZXM6IHsgX3BhcnR5aWRfdmFsdWU6IHN0cmluZyB9W10sIGlkOiBzdHJpbmcpIHtcclxuICAgIGZvciAobGV0IHBhcnR5IG9mIHBhcnRpZXMpIHtcclxuICAgICAgaWYgKHBhcnR5Ll9wYXJ0eWlkX3ZhbHVlID09IGlkKSB7XHJcbiAgICAgICAgcmV0dXJuIHBhcnR5O1xyXG4gICAgICB9XHJcbiAgICB9XHJcbiAgICByZXR1cm4gbnVsbDtcclxuICB9XHJcbn0pO1xyXG4iLCIvLyBEZW1vbnN0cmF0ZXMgdGhlIGZvbGxvd2luZyB0ZWNobmlxdWVzOlxyXG4vLyAgQ3JlYXRpbmcgcmVjb3JkcyB3aXRoIGN1c3RvbWVyIGZpZWxkcyB0aGF0IHJlZmVyZW5jZSBlaXRoZXIgYW4gYWNjb3VudCBvciBhIGNvbnRhY3RcclxuXHJcbi8vLyA8cmVmZXJlbmNlIHBhdGg9XCIuLi9XZWJBcGlSZXF1ZXN0LnRzXCIgLz5cclxuZGVzY3JpYmUoXCJcIiwgZnVuY3Rpb24oKSB7XHJcbiAgaXQoXCJDdXN0b21lciBGaWVsZHNcIiwgYXN5bmMgZnVuY3Rpb24oKSB7XHJcbiAgICB0aGlzLnRpbWVvdXQoOTAwMDApO1xyXG5cclxuICAgIGNvbnNvbGUubG9nKFwiQ3JlYXRpbmcgYWNjb3VudFwiKTtcclxuICAgIHZhciBhc3NlcnQgPSBjaGFpLmFzc2VydDtcclxuXHJcbiAgICAvLyBDcmVhdGUgYWNjb3VudFxyXG4gICAgY29uc3QgYWNjb3VudDEgPSB7XHJcbiAgICAgIG5hbWU6IGBTYW1wbGUgQWNjb3VudCAke25ldyBEYXRlKCkudG9VVENTdHJpbmcoKX1gXHJcbiAgICB9O1xyXG4gICAgdmFyIGFjY291bnQxaWQgPSAoPGFueT5hd2FpdCBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcImFjY291bnRcIiwgYWNjb3VudDEpKVxyXG4gICAgICAuaWQ7XHJcblxyXG4gICAgaWYgKCFhY2NvdW50MWlkKSB0aHJvdyBuZXcgRXJyb3IoXCJhY2NvdW50aWQgbm90IGRlZmluZWRcIik7XHJcblxyXG4gICAgLy8gQ3JlYXRlIGNvbnRhY3RcclxuICAgIGNvbnN0IGNvbnRhY3QxID0ge1xyXG4gICAgICBsYXN0bmFtZTogYFNhbXBsZSBDb250YWN0ICR7bmV3IERhdGUoKS50b1VUQ1N0cmluZygpfWBcclxuICAgIH07XHJcbiAgICB2YXIgY29udGFjdDFpZCA9ICg8YW55PmF3YWl0IFhybS5XZWJBcGkuY3JlYXRlUmVjb3JkKFwiY29udGFjdFwiLCBjb250YWN0MSkpXHJcbiAgICAgIC5pZDtcclxuXHJcbiAgICBpZiAoIWNvbnRhY3QxaWQpIHRocm93IG5ldyBFcnJvcihcImNvbnRhY3QxaWQgbm90IGRlZmluZWRcIik7XHJcblxyXG4gICAgdHJ5IHtcclxuICAgICAgLy8gQ3JlYXRlIG9wcG9ydHVuaXR5IGZvciB0aGUgY3JlYXRlZCBhY2NvdW50XHJcbiAgICAgIGNvbnN0IG9wcG9ydHVuaXR5MTogYW55ID0ge1xyXG4gICAgICAgIG5hbWU6IGBTYW1wbGUgT3Bwb3J0dW5pdHkgJHtuZXcgRGF0ZSgpLnRvVVRDU3RyaW5nKCl9YCxcclxuICAgICAgICBlc3RpbWF0ZWR2YWx1ZTogMTAwMCxcclxuICAgICAgICBlc3RpbWF0ZWRjbG9zZWRhdGU6IG5ldyBEYXRlKERhdGUubm93KCkpLnRvSVNPU3RyaW5nKCkuc3Vic3RyKDAsIDEwKSwgLy8gRGF0ZU9ubHlcclxuICAgICAgICBcImN1c3RvbWVyaWRfYWNjb3VudEBvZGF0YS5iaW5kXCI6IGBhY2NvdW50cygke2FjY291bnQxaWR9KWBcclxuICAgICAgfTtcclxuXHJcbiAgICAgIHZhciBvcHBvcnR1bml0eTFpZCA9ICg8YW55PihcclxuICAgICAgICBhd2FpdCBYcm0uV2ViQXBpLmNyZWF0ZVJlY29yZChcIm9wcG9ydHVuaXR5XCIsIG9wcG9ydHVuaXR5MSlcclxuICAgICAgKSkuaWQ7XHJcblxyXG4gICAgICBpZiAoIW9wcG9ydHVuaXR5MWlkKSB0aHJvdyBuZXcgRXJyb3IoXCJPcHBvcnR1bnR5IElEIG5vdCBkZWZpbmVkXCIpO1xyXG5cclxuICAgICAgLy8gUmV0cmlldmUgdGhlIG9wcG9ydHVuaXR5XHJcbiAgICAgIHZhciBvcHBvcnR1bml0eTIgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFxyXG4gICAgICAgIFwib3Bwb3J0dW5pdHlcIixcclxuICAgICAgICBvcHBvcnR1bml0eTFpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PW5hbWUsX2N1c3RvbWVyaWRfdmFsdWVcIlxyXG4gICAgICApO1xyXG5cclxuICAgICAgaWYgKCFvcHBvcnR1bml0eTIgfHwgIW9wcG9ydHVuaXR5Mi5fY3VzdG9tZXJpZF92YWx1ZSlcclxuICAgICAgICB0aHJvdyBuZXcgRXJyb3IoXCJPcHBvcnR1bml0eTIgQ3VzdG9tZXJpZCBub3QgcmV0dXJuZWRcIik7XHJcblxyXG4gICAgICAvLyBDaGVjayB0aGF0IHRoZSBjdXN0b21lcmlkIGZpZWxkIGlzIHBvcHVsYXRlZFxyXG4gICAgICBhc3NlcnQuaXNOb3RFbXB0eShcclxuICAgICAgICBvcHBvcnR1bml0eTIuX2N1c3RvbWVyaWRfdmFsdWUsXHJcbiAgICAgICAgXCJDdXN0b21lciBmaWVsZCBub3QgZW1wdHlcIlxyXG4gICAgICApO1xyXG5cclxuICAgICAgYXNzZXJ0LmVxdWFsKFxyXG4gICAgICAgIG9wcG9ydHVuaXR5Mi5fY3VzdG9tZXJpZF92YWx1ZSxcclxuICAgICAgICBhY2NvdW50MWlkLFxyXG4gICAgICAgIFwiQ3VzdG9tZXIgaWQgZXF1YWwgdG8gYWNjb3VudGlkXCJcclxuICAgICAgKTtcclxuXHJcbiAgICAgIGFzc2VydC5lcXVhbChcclxuICAgICAgICBvcHBvcnR1bml0eTJbXHJcbiAgICAgICAgICBcIl9jdXN0b21lcmlkX3ZhbHVlQE1pY3Jvc29mdC5EeW5hbWljcy5DUk0ubG9va3VwbG9naWNhbG5hbWVcIlxyXG4gICAgICAgIF0sXHJcbiAgICAgICAgXCJhY2NvdW50XCIsXHJcbiAgICAgICAgXCJMb2dpY2FsIG5hbWUgc2V0XCJcclxuICAgICAgKTtcclxuXHJcbiAgICAgIC8vIFVwZGF0ZSB0aGUgY3VzdG9tZXIgZmllbGQgdG8gcmVmZXJlbmNlIHRoZSBjb250YWN0XHJcbiAgICAgIGRlbGV0ZSBvcHBvcnR1bml0eTFbXCJjdXN0b21lcmlkX2FjY291bnRAb2RhdGEuYmluZFwiXTtcclxuICAgICAgb3Bwb3J0dW5pdHkxW1wiY3VzdG9tZXJpZF9jb250YWN0QG9kYXRhLmJpbmRcIl0gPSBgY29udGFjdHMoJHtjb250YWN0MWlkfSlgO1xyXG5cclxuICAgICAgYXdhaXQgWHJtLldlYkFwaS51cGRhdGVSZWNvcmQoXHJcbiAgICAgICAgXCJvcHBvcnR1bml0eVwiLFxyXG4gICAgICAgIG9wcG9ydHVuaXR5MWlkLFxyXG4gICAgICAgIG9wcG9ydHVuaXR5MVxyXG4gICAgICApO1xyXG5cclxuICAgICAgLy8gUmV0cmlldmUgdGhlIG9wcG9ydHVuaXR5XHJcbiAgICAgIHZhciBvcHBvcnR1bml0eTMgPSBhd2FpdCBYcm0uV2ViQXBpLnJldHJpZXZlUmVjb3JkKFxyXG4gICAgICAgIFwib3Bwb3J0dW5pdHlcIixcclxuICAgICAgICBvcHBvcnR1bml0eTFpZCxcclxuICAgICAgICBcIj8kc2VsZWN0PW5hbWUsX2N1c3RvbWVyaWRfdmFsdWVcIlxyXG4gICAgICApO1xyXG5cclxuICAgICAgLy8gQ2hlY2sgdGhhdCB0aGUgY3VzdG9tZXJpZCBmaWVsZCBpcyBwb3B1bGF0ZWRcclxuICAgICAgYXNzZXJ0LmlzTm90RW1wdHkoXHJcbiAgICAgICAgb3Bwb3J0dW5pdHkzLl9jdXN0b21lcmlkX3ZhbHVlLFxyXG4gICAgICAgIFwiQ3VzdG9tZXIgZmllbGQgbm90IGVtcHR5XCJcclxuICAgICAgKTtcclxuICAgICAgYXNzZXJ0LmVxdWFsKFxyXG4gICAgICAgIG9wcG9ydHVuaXR5My5fY3VzdG9tZXJpZF92YWx1ZSxcclxuICAgICAgICBjb250YWN0MWlkLFxyXG4gICAgICAgIFwiQ3VzdG9tZXIgaWQgZXF1YWwgdG8gY29udGFjdFwiXHJcbiAgICAgICk7XHJcblxyXG4gICAgICBhc3NlcnQuZXF1YWwoXHJcbiAgICAgICAgb3Bwb3J0dW5pdHkzW1xyXG4gICAgICAgICAgXCJfY3VzdG9tZXJpZF92YWx1ZUBNaWNyb3NvZnQuRHluYW1pY3MuQ1JNLmxvb2t1cGxvZ2ljYWxuYW1lXCJcclxuICAgICAgICBdLFxyXG4gICAgICAgIFwiY29udGFjdFwiLFxyXG4gICAgICAgIFwiTG9naWNhbCBuYW1lIHNldFwiXHJcbiAgICAgICk7XHJcbiAgICB9IGZpbmFsbHkge1xyXG4gICAgICAvLyBEZWxldGUgdGhlIG9wcG9ydHVuaXR5IGFuZCBhY2NvdW50IC0gb3Bwb3J0dW5pdHkgaXMgYSBjYXNjYWRlIGRlbGV0ZVxyXG4gICAgICBhd2FpdCBYcm0uV2ViQXBpLmRlbGV0ZVJlY29yZChcImNvbnRhY3RcIiwgY29udGFjdDFpZCk7XHJcbiAgICAgIGF3YWl0IFhybS5XZWJBcGkuZGVsZXRlUmVjb3JkKFwiYWNjb3VudFwiLCBhY2NvdW50MWlkKTtcclxuICAgIH1cclxuICB9KTtcclxufSk7XHJcbiJdfQ==