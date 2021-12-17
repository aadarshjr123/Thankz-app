const CosmosClient = require("@azure/cosmos").CosmosClient;
const config = require("./config");
const dbContext = require("./databaseContext");


const { endpoint, key, databaseId } = config;

const containerId="Badges"
const client = new CosmosClient({ endpoint, key,connectionPolicy: {
    enableEndpointDiscovery: false
  } });

const database = client.database(databaseId);
const container = database.container(containerId);

export async function postbadges(data){
// Make sure Tasks database is already setup. If not, create it.
await dbContext.createdab(client, databaseId, containerId);

const { resource: createdItem } = await container.items.create(data);

//console.log(`\r\nCreated new item: ${createdItem.id} - ${createdItem.description}\r\n`);
}