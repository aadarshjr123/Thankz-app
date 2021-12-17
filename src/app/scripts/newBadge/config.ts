const met = {
    endpoint: "https://thankzemp.documents.azure.com:443/",
    key: "IcHAmlrDIlcjAacJxwBuv3xy3givLfzrqNkFa0SFrBqwU1PFbWNdVlhtH8X6DaUFJselOLZa3JRj4Hv1ogyN9Q==",
    databaseId: "Thankz",
    containerId: "Badges",
    partitionKey: { kind: "Hash", paths: ["/Title"] }
  };
  module.exports = met;