'use strict';

let https = require ('https');

let host = 'westus.api.cognitive.microsoft.com';
let service = '/qnamaker/v4.0';
let method = '/knowledgebases/';

const storageName = process.env["storageName"];
const subscriptionKey = process.env["subscriptionKey"];

const { CosmosClient } = require("@azure/cosmos");
const endpoint = process.env["CosmosDBEndpoint"];
const key = process.env["CosmosDBAuthKey"];

var databaseName = process.env["DatabaseName"];
var collectionName = process.env["configCollectionName"];
const client = new CosmosClient({ endpoint, key });
const database = client.database(databaseName);
const container = database.container(collectionName);


module.exports = async function (context, myQueueItem) {
    
    //Get People KB Name
    const peopleDomain = "People";

    const querySpec = {
        query: "SELECT *  from " + collectionName + " c    WHERE c.domainName = " + "'" + peopleDomain + "'"
    };
    console.log ("Query " + JSON.stringify (querySpec));
      
    // read all items in the Items container for people department
    const { resources: items } = await container.items
    .query(querySpec)
    .fetchAll();
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: items
    };

    console.log (items[0].knowledgeBaseId);
    const kb = items[0].knowledgeBaseId;
   // Build your path URL.
    var path = service + method + kb;   
    console.log ("Path derived " + path); 
   // console.log("FileName " + context.bindingData.blobTrigger);
    var fileName = myQueueItem;
    var urlPath = storageName  + "people/" + fileName;
    console.log ("Full path after concat " + urlPath);
    var urlList = [];
    urlList.push(urlPath);
    console.log (urlList);
    let kb_model = {
        'delete': {      
          'sources': urlList          
        }
    };
      
      
    // Convert the JSON object to a string..
    let content = JSON.stringify(kb_model);
    
    // Formats and indents JSON for display.
    let pretty_print = function(s) {
        return JSON.stringify(JSON.parse(s), null, 4);
    }
    
    // Call 'callback' after we have the entire response.
    let response_handler = function (callback, response) {
        let body = '';
        response.on('data', function(d) {
            body += d;
        });
        response.on('end', function() {
        // Calls 'callback' with the status code, headers, and body of the response.
        callback ({ status : response.statusCode, headers : response.headers, body : body });
        });
        response.on('error', function(e) {
            console.log ('Error: ' + e.message);
        });
    };
    
    // HTTP response handler calls 'callback' after we have the entire response.
    let get_response_handler = function(callback) {
        // Return a function that takes an HTTP response and is closed over the specified callback.
        // This function signature is required by https.request, hence the need for the closure.
        return function(response) {
            response_handler(callback, response);
        }
    }
    
    // Calls 'callback' after we have the entire PATCH request response.
    let patch = function(path, content, callback) {
        let request_params = {
            method : 'PATCH',
            hostname : host,
            path : path,
            headers : {
                'Content-Type' : 'application/json',
                'Content-Length' : content.length,
                'Ocp-Apim-Subscription-Key' : subscriptionKey,
            }
        };
    
        // Pass the callback function to the response handler.
        let req = https.request(request_params, get_response_handler(callback));
        req.write(content);
        req.end ();
    }
    
    // Calls 'callback' after we have the response from the /knowledgebases PATCH method.
    let update_kb = function(path, req, callback) {
        console.log('host  ' + host);
        console.log('path  ' + path);
        console.log('Calling ' + host + path + '.');
        // Send the PATCH request.
        patch(path, req, function (response) {
            // Extract the data we want from the PATCH response and pass it to the callback function.
            
            callback({ operation : response.headers.location, response : response.body });
        });
    }
    
    // Calls 'callback' after we have the entire GET request response.
    let get = function(path, callback) {
        let request_params = {
            method : 'GET',
            hostname : host,
            path : path,
            headers : {
                'Ocp-Apim-Subscription-Key' : subscriptionKey,
            }
        };
    
        // Pass the callback function to the response handler.
        let req = https.request(request_params, get_response_handler(callback));
        req.end ();
    }
    let post = function (path, content, callback) {
        console.log ("In POST Function...")
        let request_params_post = {
            method : 'POST',
            hostname : host,
            path : path,
            headers : {
                'Content-Type' : 'application/json',
                'Content-Length' : content.length,
                'Ocp-Apim-Subscription-Key' : subscriptionKey,
            }
        };
        console.log ("request_params_post " + JSON.stringify (request_params_post));
    
        // Pass the callback function to the response handler.
        let reqPost = https.request (request_params_post, get_response_handler (callback));
        reqPost.write (content);
        reqPost.end ();
    }
    
    // Calls 'callback' after we have the response from the GET request to check the status.
    let check_status = function(path, callback) {
        console.log('Calling ' + host + path + '.');
        // Send the GET request.
        get(path, function (response) {
            // Extract the data we want from the GET response and pass it to the callback function.
            callback({ wait : response.headers['retry-after'], response : response.body });
        });
    }
    
    let publish_kb = function (path, req, callback) {
    
        console.log ('Calling ' + host + path + '.');
    
        // Send the POST request.
        post (path, req, function (response) {
    
            // Extract data from the POST response and pass to 'callback'.
            if (response.status == '204') {
    
                let result = {'result':'Success'};
                callback (JSON.stringify(result));
            }
            else {
                callback (response.body);
            }
        });
    }
    
    // Sends the request to update the knowledge base.
    update_kb(path, content, function (result) {
        context.log ("In update_kb function")
    
        console.log(pretty_print(result.response));
    
        // Loop until the operation is complete.
        let loop = function() {
    
            // add operation ID to the path
            path = service + result.operation;
            context.log ("path " + path)
    
            // Check the status of the operation.
            check_status(path, function(status) {
    
                // Write out the status.
                console.log(pretty_print(status.response));
    
                // Convert the status into an object and get the value of the operationState field.
                var state = (JSON.parse(status.response)).operationState;
    
                // If the operation isn't complete, wait and query again.
                if (state == 'Running' || state == 'NotStarted') {
    
                    console.log('Waiting ' + status.wait + ' seconds...');
                    setTimeout(loop, status.wait * 1000);
                } else{
                    console.log ("Publishing....")
                    var pathPost = service + method + kb; 
                    console.log ("pathPost " + pathPost)
                    publish_kb (pathPost, '', function (result) {
                        console.log (pretty_print(result));
                    });
                }
            });
        }
        // Begin the loop.
        loop();
    });
};