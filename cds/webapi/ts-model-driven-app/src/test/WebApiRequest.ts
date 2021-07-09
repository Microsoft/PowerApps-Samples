// This class is required when associating and disassociating since this is not supported by the Xrm.WebApi client side api at this time
namespace WebApiRequest {
  var webApiUrl: string = "";
  export function getWebApiUrl() {
    var context: Xrm.GlobalContext;
    var clientUrl: string;
    var apiVersion: string;
    if (webApiUrl) return webApiUrl;

    if (GetGlobalContext) {
      context = GetGlobalContext();
    } else {
      if (Xrm) {
        context = Xrm.Page.context;
      } else {
        throw new Error("Context is not available.");
      }
    }
    clientUrl = context.getClientUrl();
    var versionParts = context.getVersion().split(".");

     webApiUrl = `${clientUrl}/api/data/v${versionParts[0]}.${
      versionParts[1]
    }`;
    // Add the WebApi version
    return webApiUrl;
  }
  
  export function getOdataContext() {
    return WebApiRequest.getWebApiUrl() + "/$metadata#$ref";
  }

  export function request(
    action: "POST" | "PATCH" | "PUT" | "GET" | "DELETE",
    uri: string,
    payload?: any,
    includeFormattedValues?: boolean,
    maxPageSize?: number
  ) {
    // Construct a fully qualified URI if a relative URI is passed in.
    if (uri.charAt(0) === "/") {
      uri = WebApiRequest.getWebApiUrl() + uri;
    }

    return new Promise(function(resolve, reject) {
      var request = new XMLHttpRequest();
      request.open(action, encodeURI(uri), true);
      request.setRequestHeader("OData-MaxVersion", "4.0");
      request.setRequestHeader("OData-Version", "4.0");
      request.setRequestHeader("Accept", "application/json");
      request.setRequestHeader(
        "Content-Type",
        "application/json; charset=utf-8"
      );
      if (maxPageSize) {
        request.setRequestHeader("Prefer", "odata.maxpagesize=" + maxPageSize);
      }
      if (includeFormattedValues) {
        request.setRequestHeader(
          "Prefer",
          "odata.include-annotations=OData.Community.Display.V1.FormattedValue"
        );
      }
      request.onreadystatechange = function() {
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
              } catch (e) {
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
}
