import {Getusers,GetBadgechoice} from './cosDb_getuser';
import axios from "axios";
import * as React from "react";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
// import _ from 'lodash'
import * as microsoftTeams from "@microsoft/teams-js";
import { List, Button, Flex, Image, Text, Provider, Header, Grid, ListItem, FormatIcon, Divider, Box,Card } from '@fluentui/react-northstar'
import Container from 'react-bootstrap/Container'
import { Dropdown } from '@fluentui/react-northstar'
import { TextArea } from '@fluentui/react-northstar'
import Table from 'react-bootstrap/Table'
//import Card from 'react-bootstrap/Card'

import { postaward } from './cosmosbd_awardlog';

/**
 * State for the thankZTabTab React component
 */
export interface INewTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the thankZTabTab React component
 */
export interface INewTabTabProps {

}



  let inputItems=[{}],initialval,manager_val:any[]=[]
/**
 * Implementation of the ThankZ content page
 */
export class NewTab extends TeamsBaseComponent<any, any> {

    public constructor(props, public adapter) {
        super(props);
        this.state = {
            values:[],
            badges:[],
            status: true,
            Sendername:'',
            sImage: '',
            sName: '',
            sId: '',
            upn:undefined,
            selectedBadge: '',
            receiverName: '',
            manage:'',
            comments: '',
            Team:'',
            Upnid:'',
            curtdate:new Date().toLocaleString(),
            user:[],
            error:{}
        }
        //console.log(this.state);
    }

    GetBadge(){
        this.state.values.map((data,key) => {
            let a=<Image src={data.Picture} circular styles={{height:'3rem',padding:'0.25rem'}}/>

              this.state.badges.push({"Id":data.id,"header":data.Title,"media":a,"content":data.content})
        })
        //console.log(this.state.badges)
    }

    set_uniqueuser(){
        inputItems.shift()
        initialval.map((data,key) => {
            if(this.state.s_upn!=data.Upnid)
            inputItems.push(data)  
        })
        //console.log(inputItems)

    }

    assignmanager(tempmanager){
        //console.log(tempmanager)
        manager_val.pop()
        console.table(initialval)
        initialval.map((data,key) => {
            if(tempmanager==data.Upnid)
            manager_val.push(data)  
           
            
        })
        
        //console.log(manager_val)
    }

    

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        initialval=await Getusers()
        //console.log(this.state.login_name)
        let lists = await GetBadgechoice()
        await this.setState({
            values:lists
        });
        
        this.GetBadge();       
        
        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                
                this.setState({
                    entityId: context.entityId,
                    s_upn:context.userPrincipalName,
                    
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
    

       

        const submitHandler = () => {
            this.setState({
                status: true,
                receiverName: '',
                comments: ''
            })
        }

        const handle1 = (index) => {
           this.set_uniqueuser()
            const value = this.state.badges[index];
            this.setState({
                sImage: value.media.props.src,
                sName: value.header,
                status: false
            })
        }

     

        const handleChange = (event, value) => {
            //console.log(value.value)
            
            this.assignmanager(value.value.manager)
            this.setState({
                receiverName: value.value.header,
                profile:value.value.profile,
                receivers:value.value,
                Team:value.value.Team,
                Upnid:value.value.Upnid,
                
            })
            
            

        }

        const handleChangeManager = (event, value) => {
            //console.log(value)
            //console.log(value.value.header)
            //console.log(event)
            this.setState({
                managers:value.value,
                manage:value.value.header
            })

        }

        const handleChangePeers = (event, value) => {
            
            //console.log(event)
            this.setState({
                peers:value.value
            })

        }

        const getA11ySelectionMessage = {
            onAdd: item => `${item.header} has been removed.`,
            onRemove: item => `${item.header} has been removed.`,
        }

        const handleTextChange = (event, value) => {
           
            this.setState({
                comments: value.value
            })
        }
        const validation=()=>{
            var err={
                receiver:"",
                manager:"",
                comment:""
            }
            var valid=true

            if(!this.state.receiverName){
                valid=false
                err.receiver="*required"
            }

            if(!this.state.manage)
            {
                valid=false
                err.manager="*required"
            }

            if(!this.state.comments)
            {
                valid=false
                err.comment="*required"
            }

            this.setState({
                error:err
            })

            return valid

        }
      

        const callHandler = async() => {
            microsoftTeams.initialize();
            let isvalid=await validation()
            if(isvalid){
            let username;
            for(let a of initialval){
                if(a.Upnid==this.state.s_upn)
                    username=a.header
                    
                }

                let newdata={"name":username,"Badges":this.state.sName,"Image":this.state.sImage,"Receivers":this.state.receivers,"managers":this.state.managers,"peers":this.state.peers,"comment":this.state.comments}
            microsoftTeams.tasks.submitTask(newdata)
            let data = {"Sender":username,"Badges":this.state.sName,"Image":this.state.sImage,"ReceivedTo":this.state.receiverName,"ReceivedAt":this.state.curtdate,"ReceiverEmail":this.state.Upnid,"SenderMail": this.state.s_upn,"ReceiverTeam":this.state.Team,"comments":this.state.comments,profile:this.state.profile};
        
            postaward(data)
           
        }
    }

      
            

        return (
            <Provider theme={this.state.theme} style={{maxWidth:'100vw'}}>
                

                { 
                    (this.state.status === true )
                    ? <div > 
                        {/* {nameList}  */}

                        <h4 style={{fontFamily:'Segoe UI',fontSize:'15px'}} >Select a badge</h4>
                  
                   
                         <div style={{marginLeft:'1rem'}} >
                         <List selectable items={this.state.badges} 
                            horizontal
                            styles={{display:'grid',gridTemplateColumns: "repeat(auto-fill, minmax(250px, 1fr))", gridGap:'2rem 2rem'}}
                            className="badgesrow badges li"
                            defaultSelectedIndex={0}
                            selectedIndex={this.state.selectedBadge}
                            onSelectedIndexChange={(e, newProps: any) => {
                                handle1(newProps.selectedIndex)
                                //console.log(this.state)
                            }} />
                            </div>           
                  
                   
                    </div> 
                    : <div style={{display:'grid',gridTemplateColumns: "repeat(auto-fill, minmax(350px, 1fr))",gridGap:'20px 20px' }}>

                        <div>
                            <div>
                            <h5 style={{ marginTop: '30px' }}>Receiver<sup style={{color:'red'}}>{this.state.error.receiver}</sup></h5>
                            <Dropdown
                                fluid
                                clearable
                                style = {{ height: '40px',maxWidth:'425px'}}
                                items={inputItems}
                                placeholder="Receiver"
                                onChange={handleChange}
                                getA11ySelectionMessage={getA11ySelectionMessage}
                                
                            />
                            </div>
                            <div>
                            <h5 className="mt-2">Manager<sup style={{color:'red'}}>{this.state.error.manager}</sup></h5>
                            <Dropdown
                                fluid
                                clearable
                                style = {{maxWidth:'425px'}}
                                items={manager_val}
                                onChange={handleChangeManager}
                                placeholder="Managers"
                                getA11ySelectionMessage={getA11ySelectionMessage}
                                noResultsMessage="We couldn't find any matches."
                                
                            />
                            </div>
                            <div>                        
                            <h5 className="mt-2">Peers(Optional)</h5>
                            <Dropdown
                                fluid
                                multiple
                                search
                                style = {{maxWidth:'425px'}}
                                items={inputItems}
                                onChange={handleChangePeers}
                                placeholder="Peers"
                                getA11ySelectionMessage={getA11ySelectionMessage}
                                noResultsMessage="We couldn't find any matches."
                            />
                            </div>
                            <h5 className="mt-2">Comments <sup style={{color:'red'}}>{this.state.error.comment}</sup></h5>
                            <div>
                                <TextArea fluid onChange={handleTextChange} styles={{ marginTop: '5px',maxWidth:'425px'}} resize="both" placeholder="Add comments..." />
                        
                            </div>
                        </div>  
                        <div>  
                            <Card fluid style={{maxWidth:'25rem',
                                        maxHeight:'27rem',
                                        border:'1px',
                                        borderTop:'3px Solid',
                                        borderStyle:'solid',
                                        borderRadius:'5px',
                                        borderColor:'#D2D2D0',
                                        borderTopColor:'Red'}}>
                            <Card.Header>
                            
                                <Flex style={{ margin: '0 0 0 10px'}} >
                                    
                                    <Text style={{ marginLeft: '10px'}} content="To:" weight="bold" />
                                    
                                    <Text style={{ marginLeft: '5px',minWidth: '200px',}} content={this.state.receiverName}  />
                                    
                                    
                            </Flex>
                                
                            </Card.Header> 
                            <Card.TopControls>
                                <Flex>
                                    <Box
                                    content="Box"
                                    styles={{
                                    border: '2px solid #ccc',
                                    width: '120px',
                                    borderRadius:'3px',
                                    marginRight:'10px',
                                    textAlign:'center'
                                }}
                                
                                >
                                    <Image
                                    circular
                                    styles={{margin:'5px',width:'1rem'}}
                                    src={this.state.sImage} />
                                    <Text  content={ this.state.sName } />
                                    </Box>
                                </Flex>
                                </Card.TopControls>        
                            <Divider size={1} color={'orange'} style={{ margin: '10px'}}/>
                            <Card.Footer>
                                <div style={{ backgroundColor:'#FFFADE',margin:'20px',borderRadius:'5px'}}>
                                    <Image fluid
                                    className="col-xs-1 text-center"
                                styles={{width: '8rem',position:'relative',paddingTop:'1rem',paddingBottom:'1rem',margin:'auto',display:'block'}}
                                    src={this.state.sImage} />
                                </div>
                                <Text style={{ marginLeft: '15px'}} content="Comment:"  />
                                <div style={{height:"120px",overflow:"scroll"}}>
                                <Text style={{ marginLeft: '20px'}} content={ this.state.comments }  />
                                </div>
                            </Card.Footer>
                            
                            </Card>

                                {/* <img style={{ height: '8rem', marginBottom: '150px'}} src={this.state.sImage} /> */}
                                <div style={{ margin: '10px'}}>
                                
                                                <Button content="Primary" style={{ margin: '5px'}} onClick={() =>  submitHandler() }>Back</Button>
                                            
                                                <Button primary content="Secondary" style={{ margin: '5px'}} onClick={() => callHandler()}>Send</Button>
                                </div>         
                        </div>
                    
                         
                    </div> 
                }
                {/* {nameList} */}
                
            </Provider>
        );
    }
}
