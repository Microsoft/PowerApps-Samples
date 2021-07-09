// Demonstrates the following techniques:
//  Using the Bound function AddToQueueResponse
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-actions#bound-actions

describe("", function() {

  it("AddToQueueResponse", async function() {
    this.timeout(90000);
    var assert = chai.assert;
    // Create Queue
    var queueid = (<any>(await Xrm.WebApi.createRecord("queue",{"name" : "Sample Queue"}))).id;
    
    // Create letter
    var letterid = (<any>(await Xrm.WebApi.createRecord("letter",{"subject" : "Sample Letter"}))).id;

    try{
    // Execute request
    var AddToQueueRequest = new class {
      entity = {
        id: queueid,
        entityType: "queue"
      };
      Target = {
          id: letterid,
          entityType: "letter"
      };

      getMetadata(): any {
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
      }
    }();


    var response : {
      QueueItemId: string
    }= await (<any>(await Xrm.WebApi.online.execute(AddToQueueRequest))).json();
  
    assert.isString(response.QueueItemId,"QueueItemId returned");

    }
    finally{
      // Delete Letter
      if (letterid) {
        await Xrm.WebApi.deleteRecord("letter", letterid);
      }
       // Delete Queue
       if (queueid) {
        await Xrm.WebApi.deleteRecord("queue", queueid);
      }
    }
    

  });
});
