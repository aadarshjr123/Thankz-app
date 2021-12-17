const axios = require("axios");

export async function GetAccessToken() {
  let a = axios
    .get("https://https://quadrathankz.azurewebsites.net/token")
    .then(function (response) {
      
      return response.data.accessToken;
    })
    .catch(function (error) {
      //console.log(error);
    });
  return a;
}
