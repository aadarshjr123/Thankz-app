import { GetAccessToken } from "./token";
const req=require("request").defaults({ encoding: null });
const axios = require("axios");

export async function GetProfilePicture(upn) {
     let token =await GetAccessToken()
    const graphEndpoint = "https://graph.microsoft.com/v1.0/users/"+upn+"/photo/$value";

    const response = await axios(graphEndpoint, { headers: { Authorization: `Bearer ${token}` }, responseType: 'arraybuffer' });
    const avatar ="data:" + response.headers["content-type"] + ";base64," + new Buffer(response.data).toString('base64');
    
    return avatar
// var imagedata
// let profile=req.get({
//     url:"https://graph.microsoft.com/v1.0/users/"+upn+"/photo/$value",
//     headers: {
//       "Authorization": "Bearer " + token
//     }
//   }, async function(err, response, body) {
//     if (!err && response.statusCode == 200) {
       
//       imagedata = "data:" + response.headers["content-type"] + ";base64," + new Buffer(body).toString('base64');
//       //console.log(imagedata);
      
      
//   }
   
//   })
//   let a=await imagedata
//   //console.log(a)
}