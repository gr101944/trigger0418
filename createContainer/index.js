const { BlobServiceClient } = require('@azure/storage-blob');
const AZURE_STORAGE_CONNECTION_STRING =  process.env.AZURE_STORAGE_CONNECTION_STRING;
if (!AZURE_STORAGE_CONNECTION_STRING) {
    throw Error("Azure Storage Connection string not found");
}

module.exports = async function (context, req) {
    console.log (JSON.stringify(req))
    console.log (JSON.stringify(req.query.containerName))
    const blobServiceClient = BlobServiceClient.fromConnectionString(
        AZURE_STORAGE_CONNECTION_STRING
    );
    
    // Create a unique name for the container
    const containerName = req.query.containerName;
    
    console.log("\nCreating container...");
    console.log("\t", containerName);
    
    // Get a reference to a container
    const containerClient = blobServiceClient.getContainerClient(containerName);
    // Create the container
    const createContainerResponse = await containerClient.create();
    console.log(
    "Container was created successfully. requestId: ",
    createContainerResponse.requestId
    );
}