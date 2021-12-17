import * as React from "react";
// import { Provider, Flex, Text, Button, Header,FormatIcon,Divider} from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { List, Button, Flex, Image, Text, Card, Provider, Grid,FormatIcon, Divider } from '@fluentui/react-northstar'
import { Input } from '@fluentui/react-northstar'

import { postbadges } from "./cosmosbd_uploadbadge";


export class NewBadge extends TeamsBaseComponent<any, any> {

    constructor(props){
        super(props)
        this.state = {
          file: null,
          temp:null,
          uri:"",
          title: '',
          description: '',
          error:{}
        }
    }

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    public render() {
        
        // let datas = Sample(data);
        // //console.log(datas);
        const handleChange = (event) => {
            let fil=event.target.files[0]
            this.setState({
                temp: URL.createObjectURL(event.target.files[0]),
                file:fil
            })
        }

        const handleTitleChange = (event) => {
            this.setState({
                title: event.target.value
            })
        }

        const handleDescriptionChange = (event) => {
            this.setState({
                description: event.target.value
            })
        }
        const validation=()=>{
            var err={
                file:"",
                Title:"",
                Description:""
            }
            var valid=true

            if(!this.state.file){
                valid=false
                err.file="*required"
            }

            if(!this.state.title)
            {
                valid=false
                err.Title="*required"
            }

            if(!this.state.description)
            {
                valid=false
                err.Description="*required"
            }

            this.setState({
                error:err
            })

            return valid

        }

        const create = async() => {
            microsoftTeams.initialize();
           // let token = await GetAccessToken();
           let isvalid=validation()
if(isvalid){
            var fil = this.state.file;
            let a=this.state.title
            let b=this.state.description
            let c=this.state.uri
            var reader = new FileReader();
            reader.onloadend =async function() {
                c=reader.result
            let data = {"Title":a,"content":b,"Picture":c};
            
             var datas = JSON.stringify(data);
             await postbadges(data)
            }
            reader.readAsDataURL(fil);
        setTimeout(()=> microsoftTeams.tasks.submitTask(),1700);
            }
        }

       


        return (

            <Provider theme={this.state.theme} >
    
    <div style={{display:'grid',gridTemplateColumns: "repeat(auto-fill, minmax(350px, 1fr))",gridGap:'20px 20px' }}>
    
      <div style={{maxWidth:'380px',marginTop:'30px'}}>
        <div>
          <label >Badge title<sup style={{color:'red'}}>{this.state.error.Title}</sup></label>
          <br/>
          <Input fluid placeholder="Add name" onChange = { handleTitleChange } styles={{ marginTop: '15px',maxWidth:'380px'}} />
        </div>
        <br/>
        <div>                                  
          <label >Description<sup style={{color:'red'}}>{this.state.error.Description}</sup></label>                                 
          <br/>                                    
          <Input fluid maxLength={40} type="text" placeholder="Add text" onChange = { handleDescriptionChange } styles={{ marginTop: '15px', maxWidth: '380px' }} />
        </div>
        <br/>
        <div>
          <label >Add Icon Image<sup style={{color:'red'}}>{this.state.error.file}</sup></label>
          <Card fluid style={{ maxWidth: '380px',height: '50px', border: '1px dashed', marginTop: '15px'}}>
              <Card.Footer className="col-xs-1 text-center">
                  <input type="file" onChange={handleChange} ></input>
              </Card.Footer>
          </Card>
        </div>                 
        </div> 

                                             
  <div >
  <Card fluid style={{maxWidth:'34rem',
                    maxHeight:'25rem',
                    minWidth:'17rem',
                    border:'1px',
                    borderTop:'3px Solid',
                    borderStyle:'solid',
                    borderRadius:'5px',
                    borderColor:'#D2D2D0',
                    borderTopColor:'Red'}}>
                        <Card.Header >
                            <div style={{ backgroundColor:'#FFFADE',margin:'20px',borderRadius:'5px'}} className="col-xs-1 text-center" >
                                <Image className="col-xs-1 text-center"
                                styles={{ height: '10rem', width: '10rem',position:'relative',margin:'auto',marginTop:'1rem',marginBottom:'1rem',display:'block'}}
                                src={this.state.temp} />
                            </div>
                        </Card.Header>
                        <Divider size={1} color={'orange'} />
                        <Card.Footer>
                            <Flex style={{ margin: '5px 0 10px 5px'}} >
                                <FormatIcon/>
                                <Text style={{ marginLeft: '10px'}} content=" Badge Title" weight="bold" />
                                <Text style={{ marginLeft: '30px'}} content={ this.state.title }  weight="bold" />
                            </Flex>
                            <Flex style={{ margin: '5px'}} >
                                <FormatIcon />
                                <Text style={{ marginLeft: '10px'}} content=" Description" weight="bold" />
                                <Text style={{ marginLeft: '20px',marginRight:'2px'}} content={ this.state.description }></Text>
                            </Flex>
                        </Card.Footer>         
                    </Card>
                    <Button primary style={{ margin: '10px'}} onClick={() => create()}>Create</Button>
  </div>
  </div>
</Provider>
       
        )

    }
}