const CosmosClient = require("@azure/cosmos").CosmosClient;
const config = require("../cosmosdb/config");
const dbContext = require("../cosmosdb/databaseContext");


const { endpoint, key, databaseId, containerId } = config;

const client = new CosmosClient({ endpoint, key,connectionPolicy: {
    enableEndpointDiscovery: false
  } });

const database = client.database(databaseId);
const container = database.container(containerId);

export async function postuser(data,upn){
// Make sure Tasks database is already setup. If not, create it.
await dbContext.create(client, databaseId, containerId);
try{
const { resource: createdItem } = await container.items.create(data);

//console.log(`\r\nCreated new item: ${createdItem.id} - ${createdItem.description}\r\n`);
}
catch(err)
{
    //console.log(err)
}

}