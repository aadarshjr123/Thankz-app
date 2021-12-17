const axios = require("axios");

export async function GetAccessToken() {
  let a = axios
    .get("https://3acf-59-92-24-114.ngrok.io/token")
    .then(function (response) {
      
      return response.data.accessToken;
    })
    .catch(function (error) {
      //console.log(error);
    });
  return a;
}
