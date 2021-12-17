import { itemLayoutSlotClassNames } from "@fluentui/react-northstar";

const Cosmos= require("@azure/cosmos").CosmosClient;
const { endpoint, key, databaseId, containerIduser,containerIdbadge,containerIdaward} = {
    endpoint: "https://thankzemp.documents.azure.com:443/",
    key: "IcHAmlrDIlcjAacJxwBuv3xy3givLfzrqNkFa0SFrBqwU1PFbWNdVlhtH8X6DaUFJselOLZa3JRj4Hv1ogyN9Q==",
    databaseId: "Thankz",
    containerIduser: "Users",
    containerIdbadge: "Badges",
    containerIdaward: "Awardlog"
  };

let userid

const client = new Cosmos({ endpoint, key,connectionPolicy: {
  enableEndpointDiscovery: false
} });
let details,badges,award,list

const database = client.database(databaseId);
const containeruser = database.container(containerIduser);//Userz
const containerbadge = database.container(containerIdbadge);//Badges
const containeraward = database.container(containerIdaward);//Awardlog
export async function Getusers() {
  try{
 const querySpec = {
  query: "SELECT * from c"
};

// read all items in the Items container
const { resources: items } = await containeruser.items
  .query(querySpec)
  .fetchAll();

details=items
//console.log(items)
  }
  catch(err)
  {
    //console.log(err)
  }

return details
}

export async function GetBadgechoice() {
  try{
 const querySpec = {
  query: "SELECT * from c"
};

// read all items in the Items container
const { resources: items } = await containerbadge.items
  .query(querySpec)
  .fetchAll();

badges=items
//console.log(items)
  }
  catch(err)
  {
    //console.log(err)
  }

return badges
}

export async function Getawardlog() {
  if(award!=undefined){
    //console.log(award)
  }
  else{
    try{
      const querySpec = {
      query: "SELECT * from c"
  };

  // read all items in the Items container
  const { resources: items } = await containeraward.items
    .query(querySpec)
    .fetchAll();

  award=items
  //console.log(items)
    }
    catch(err)
    {
      //console.log(err)
    }
  }
return award
}

export async function Getuniqueuser(upn) {
//console.log(upn)
userid=upn
  try{
    const querySpec = {
     query: `SELECT * from c where c.Upnid="${upn}"`
   };
   
   // read all items in the Items container
   const { resources: items } = await containeruser.items
     .query(querySpec)
     .fetchAll();
   
   
   //console.log(items.length)

   return items.length
     }
     catch(err)
     {
       //console.log(err)
     }
   
}
