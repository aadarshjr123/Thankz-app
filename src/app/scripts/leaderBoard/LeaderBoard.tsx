import * as React from "react";
import { Provider,Image} from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

import Table from 'react-bootstrap/Table'


import { Getawardlog } from "../newTab/cosDb_getuser";

/**
 * State for the thankZTabTab React component
 */

export interface IThankZTabStates extends ITeamsBaseComponentState {
  entityId?: string;
  upn?: string;
}

/**
* Properties for the thankZTabTab React component
*/
export interface IThankZTabPropss {
}


const listview={
  paddingLeft:"8vw"
  };
  
/**
 * Implementation of the ThankZ content page
 */
export class LeaderBoard extends TeamsBaseComponent<any,any> {
  public constructor(props, public adapter) {
    super(props);
    this.state = {
        values: {},
        some:[],
        Team:[],
        date:new Date(),
        tenantID: process.env.MY_TenantID,
        userTenantID: ""
              } 
        
}

  most_appreciated()
  {
    let obj={}
    this.state.values.map(async(data,key) => {
      if(data.ReceiverEmail in obj)
        obj[data.ReceiverEmail].Total_badges++
      else
      
        obj[data.ReceiverEmail]={'name':data.ReceivedTo,'receiver_team':data.ReceiverTeam,'Total_badges':1,'Email':data.ReceiverEmail,'profile':data.profile}
    })   

    Object.values(obj).map((val)=>{
      this.state.some.push(val)
    })

    this.state.some.sort(function(a,b){
      return b.Total_badges-a.Total_badges
    })

    console.table(this.state.some)
  }




  Lists = () => (

    this.state.some.slice(0,6).map((data,Key) => {
      
      return (
        <tr>
        <td style={listview}>
          <div>
          <Image fluid circular src={data.profile} 
                               style={{
                                height: "1.80rem",
                                width: "1.80rem"
                              }} />
              <span style={{marginLeft:"5px"}}>{data.name}</span>
          </div>
        </td>
        <td style={listview}>{data.receiver_team}</td>
        <td style={listview}>{data.Total_badges}</td>
      </tr>
      )
    })
  )



  Team_Token(){
    
    let obj={};
    this.state.values.map((data,key) => {
      
        if(data.ReceiverTeam in obj) {
          obj[data.ReceiverTeam].Total_badges++;      
        
        }
        else{
          obj[data.ReceiverTeam]={'receiver_team':data.ReceiverTeam,'Total_badges':1}
        
          
        }
    })
    console.table(obj)
      
    Object.values(obj).map((val)=>{
      this.state.Team.push(val)

      this.state.Team.sort(function(a,b){
        return b.Total_badges-a.Total_badges
      })
      
    
    })

      
  }

  Teamlists = () => (
    
    this.state.Team.map((mail,key) => {
      
      return (
        <tr>
        <td style={listview}>{mail.receiver_team}</td>
        <td style={listview}>{mail.Total_badges}</td>
        </tr>
                  )
        }) 
  )

    
  public async componentWillMount() {
      this.updateTheme(this.getQueryVariable("theme"));
      let lists = await Getawardlog();
      
      await this.setState({
          values: lists,
      });
      this.most_appreciated();
      this.Team_Token();

      if (await this.inTeams()) {
          microsoftTeams.initialize();
          microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
          microsoftTeams.getContext((context) => {
              //console.log(context)
              microsoftTeams.appInitialization.notifySuccess();
              this.setState({
                  entityId: context.entityId,
                  upn: context.userPrincipalName,
                  userTenantID: context.tid

              });
              this.updateTheme(context.theme);
          });
      } else {
          this.setState({
              entityId: "This is not hosted in Microsoft Teams"
          });
      }
  }

    

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
      if(this.state.tenantID === this.state.userTenantID) {
        return (
          <Provider theme={this.state.theme} style={{maxWidth:'93vw',backgroundColor:"white"}}>

          <div>
          <h5 style={{ margin:"40px 0 10px 0",color:"#292929"}}>Most Appreciated Users</h5> 
          <div >
            <Table striped hover >
              <thead>
                <tr>
                  <th style={listview}>Receiver Name</th>
                  <th style={listview}>Receiver Team</th>
                  <th style={listview}>Total Number Of badges</th>
                </tr>
              </thead>
              <tbody>
              <this.Lists/>
              </tbody>
            </Table>
            </div>

            <h5 style={{ margin:"70px 0 10px 0",color:"#292929"}}>Team on top</h5> 
            <div >
            <Table striped hover >
              <thead>
                <tr>
                  <th style={listview}>Team </th>
                  <th style={listview}>Total Number Of badges</th>
                </tr>
              </thead>
              <tbody>
              <this.Teamlists />
              </tbody>
            </Table>
            </div>
            </div>
        </Provider>
             
        )
      } else {
        return (
          <Provider theme={this.state.theme} style={{ maxWidth: '93vw', backgroundColor: "white" }} >
           <div style={{textAlign: "center",marginTop: "10%",position: "relative",width: "100%"}}>
              <img style={{maxWidth:"100%",maxHeight:"100%"}} src="https://res.cloudinary.com/dgrljsghp/image/upload/v1639719110/Illustration_o4d0qu.svg" data-reactid=".0.0"/>
              <div style={{margin: "20px"}}/>
              <p style={{fontWeight: "bold"}}>Welcome to Quadra Thankz!</p>
              <p style={{fontWeight: "lighter"}}>Please reach out sales/marketing team of Quadrasystems.Net India pvt ltd at css@quadrasystems.net  if your organization is intrested to use Thankz app along with Quadrasystems</p>
              </div>
        </Provider>
      );
      }
        
        
    }
}
