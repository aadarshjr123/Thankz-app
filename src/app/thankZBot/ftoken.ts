const request = require("request");
const req=require("request").defaults({ encoding: null });
import { Getuniqueuser } from '../scripts/newTab/cosDb_getuser';
import { postuser } from './cosDb_user';
export async function GetuserToken(upn,data){

  
  
if(!await Getuniqueuser(upn))

{
const endpoint = "https://login.microsoftonline.com/08948d7c-43ee-4cae-9f2c-67e0464345d8/oauth2/v2.0/token";
const requestParams = {
    grant_type: "client_credentials",
    client_id: "a013c4a4-0683-4125-8c05-4004c2c3cc6f",
    client_secret: "5CVaG3r.a3y_OK2K9PRa39hy~lqc-HQC3a",
    scope: "https://graph.microsoft.com/.default"
};
var aa,detail,team
request.post({ url: endpoint, form: requestParams }, function (err, response, body) {
    if (err) {
        //console.log("error");
    }
    else {
        
        
        let parsedBody = JSON.parse(body);
        if (parsedBody.error_description) {
            //console.log("Error=" + parsedBody.error_description);
        }
        else {
            
            //console.log("Access Token=" + parsedBody.access_token);
            aa=parsedBody.access_token
            request.get({
              url:"https://graph.microsoft.com/beta/users/"+upn+"/profile/positions",
              headers: {
                "Authorization": "Bearer " + aa
              }
            }, async function(err, response, body) {
              ////console.log(body);
              
              let to_obj=JSON.parse(body)
              //console.log(to_obj)
              //console.log(to_obj["value"][0].detail.company.department)
              team=to_obj["value"][0].detail.company.department
              data.Team=team
              
            


            request.get({
                url:"https://graph.microsoft.com/v1.0/users/"+upn+"/",
                headers: {
                  "Authorization": "Bearer " + aa
                }
              }, async function(err, response, body) {
                  //console.log("sdkjasnkdj")
                //console.log(body);
                let to_obj=JSON.parse(body)
                data.header=to_obj.displayName
                //await postuser(data,upn)
              
              
              request.get({
                url:"https://graph.microsoft.com/v1.0/users/"+upn+"/manager",
                headers: {
                  "Authorization": "Bearer " + aa
                }
              }, async function(err, response, body) {
                  //console.log("sdkjasnkdj")
                //console.log(body);
                let to_obj=JSON.parse(body)
                data.manager=to_obj.userPrincipalName
                //await postuser(data,upn)

                req.get({
                  url:"https://graph.microsoft.com/v1.0/users/"+upn+"/photo/$value",
                  headers: {
                    "Authorization": "Bearer " + aa
                  }
                }, async function(err, response, body) {
                  if (!err && response.statusCode == 200) {
                    let imagedata = "data:" + response.headers["content-type"] + ";base64," + new Buffer(body).toString('base64');
                    //console.log(imagedata);
                    data.profile=imagedata
                    await postuser(data,upn)
                }
                  
                })

              })
            })
          })                
        
    }
}
}) 
//console.log("logan")
return detail
}
}
