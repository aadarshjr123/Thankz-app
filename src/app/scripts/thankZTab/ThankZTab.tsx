import * as React from "react";
import { Provider, Flex, Segment, Image, Text, Button, Header, Card, CardBody, CardHeader, Avatar, Layout, ArrowUpIcon, ArrowDownIcon } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

import Table from 'react-bootstrap/Table'
import Container from 'react-bootstrap/Container'

import "./styles.css";
import Row from 'react-bootstrap/Row'
import Col from 'react-bootstrap/Col'
import { Getawardlog } from "../newTab/cosDb_getuser";
/**
 * State for the thankZTabTab React component
 */

export interface IThankZTabState extends ITeamsBaseComponentState {
  entityId?: string;
  upn?: string;
}

/**
 * Properties for the thankZTabTab React component
 */
export interface IThankZTabProps {

}


const listview = {
  paddingLeft: "8vw"
};

/**
 * Implementation of the ThankZ content page
 */
export class ThankZTab extends TeamsBaseComponent<any, any> {
  public constructor(props, public adapter) {
    super(props);
    this.state = {
      values: [],
      list: [],
      some: [{}],
      sentto: [{}],
      upn: undefined,
      sent: true,
      recev: false,
      tenantID: process.env.MY_TenantID,
      userTenantID: ""
    }
    // //console.log(this.state);
  }



  ListToArray() {
    // let some = [{}]
    let upi
    microsoftTeams.initialize();
    microsoftTeams.getContext((context) => {

      //console.log(context);
      upi = context.userPrincipalName

      this.setState({
        userTenantID: context.tid
      })

      this.state.values.map((data, key) => {

        if (upi === data.ReceiverEmail) {
          this.state.some.push({ "Badges": data.Badges, "ReceivedTo": data.ReceivedTo, "ReceivedAt": data.ReceivedAt, "Image": data.Image, "Sender": data.Sender })

        }

        if (data.SenderMail === upi) {
          this.state.sentto.push({ "Badges": data.Badges, "SentTo": data.ReceivedTo, "SentAt": data.ReceivedAt, "Image": data.Image, "Sender": data.Sender })

        }


      })
      //console.log(upi)
    });
    //console.log(this.state.some)
    //console.log(this.state.sentto)

    return "finished"
  }



  senderCount = () => {
    let count = 0;
    this.state.values.map((data, key) => {
      if (data.SenderMail === this.state.upn) {
        return count++;
      }
    })

    return count;
  }


  receiverCount = () => {
    let count = 0;
    this.state.values.map((data, key) => {
      if (data.ReceiverEmail === this.state.upn) {
        return count++;
      }
    })
    return count;
  }

  Lists = () => (
    this.state.some.slice(1).map((data, key) => {

      return (
        <tr style={{ marginLeft: '2rem' }}>
          <td style={listview}>
            <div>
              <Image circular fluid src={data.Image}
                style={{
                  height: "1.80rem",
                  width: "1.80rem"
                }} />{data.Badges}
            </div>
          </td>
          <td style={listview}>
            <div>
              {data.Sender}
            </div>
          </td >
          <td style={listview}>{data.ReceivedAt}</td>
        </tr>
      )
    })
  )

  Sentitem = () => (
    this.state.sentto.slice(1).map((data, key) => {

      return (
        <tr style={{ marginLeft: '2rem' }} >
          <td style={listview}>
            <div>
              <Image circular fluid src={data.Image}
                style={{
                  height: "1.80rem",
                  width: "1.80rem"
                }} />{"  " + data.Badges}
            </div>
          </td>
          <td style={listview}>
            <div>
              {data.SentTo}
            </div>
          </td>
          <td style={listview}>
            {data.SentAt}
          </td>
        </tr>
      )
    })
  )
  sentdis = () => {
    this.setState({
      sent: true,
      recev: false
    })
  }
  recevdis = () => {
    this.setState({
      sent: false,
      recev: true
    })
  }

  CardExample2 = () => (
    <Card styles={{
      maxWidth: '700px'
    }} fluid aria-roledescription="card with avatar, image and action buttons" onClick={() => this.sentdis()}>
      <Card.Header>
        <Flex gap="gap.small">
          <ArrowUpIcon style={{ margin: '5px' }} />
          <Flex column style={{ marginLeft: '10px' }}>
            <Text content={this.senderCount()} weight="bold" />
            <Text content="Appreciation Sent" size="small" />
          </Flex>
        </Flex>
      </Card.Header>

    </Card>
  )

  CardExample = () => (
    <Card styles={{
      maxWidth: '700px',
    }} fluid aria-roledescription="card with avatar, image and action buttons" onClick={() => this.recevdis()}>
      <Card.Header>
        <Flex gap="gap.small">
          <ArrowDownIcon style={{ margin: '5px' }} />
          <Flex column style={{ marginLeft: '10px' }}>
            <Text content={this.receiverCount()} weight="bold" />
            <Text content="Appreciation Received" size="small" />
          </Flex>
        </Flex>
      </Card.Header>

    </Card>
  )

  public async componentWillMount() {

    this.updateTheme(this.getQueryVariable("theme"));
    let lists = await Getawardlog();
    //console.log(lists);
    await this.setState({
      values: lists,
    });
    let a = await this.ListToArray();
    //console.log(a)


    if (await this.inTeams()) {
      microsoftTeams.initialize();
      microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

      microsoftTeams.getContext(async (context) => {

        //console.log(context);
        microsoftTeams.appInitialization.notifySuccess();
        this.setState({
          entityId: context.entityId,
          upn: context.userPrincipalName
        });




        //this.updateTheme(context.theme);

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

    const submit = () => {
      //console.log('entered');
      microsoftTeams.initialize();
      let taskInfo = {
        title: "Add Badges",
        height: 550,
        width: 950,
        url: "https://https://quadrathankz.azurewebsites.net/newBadge",
        completionBotId: process.env.MICROSOFT_APP_ID
      };
      let submitHandler = (err, result) => {
        //console.log(result, err);
      };
      microsoftTeams.tasks.startTask(taskInfo, submitHandler);
      microsoftTeams.tasks.submitTask();
    }


    // microsoftTeams.tasks.startTask(taskInfo, submitHandler);

    // const openTaskModule = () => {

    //   microsoftTeams.tasks.startTask(taskInfo);
    // }
    //className="col-xs-1 text-center"
    //{this.state.tenantID === this.state.userTenantID}
  
    if(this.state.tenantID === this.state.userTenantID) {
      return (

        <Provider theme={this.state.theme} style={{ maxWidth: '93vw', backgroundColor: "white" }} >
          <br />
          {this.state.upn === process.env.adminacc ?
  
            <Container>
              <Row>
  
                <Col style={{ float: 'right' }}>
                  <Text content="Create your organisation's specific badges   " style={{ color: "#292929" }} /> <Button className="m-1" onClick={() => submit()}>Create Badge</Button>
                </Col>
              </Row>
            </Container>
            :
            <div>
  
            </div>
          }
  
          <br />
            <div>
              <Container>
                <Row>
                  <Col>
                    <this.CardExample />
                  </Col>
                  <Col>
                    <this.CardExample2 />
                  </Col>
                </Row>
              </Container>
              <br />
              <div>
                {(this.state.recev) ?
                  (
                    <div >
                      <Table striped hover className="listview" >
                        <thead>
                          <tr>
                            <th style={listview}>Badges</th>
                            <th style={listview}>Received From</th>
                            <th style={listview}>Received At</th>
                          </tr>
                        </thead>
                        <tbody>
                          <this.Lists />
                        </tbody>
                      </Table>
                    </div>
                  ) : ""}
  
              </div>
              <div>
                {(this.state.sent) ?
                  (
                    <div >
                      <Table striped hover className="listview" >
                        <thead>
                          <tr >
                            <th style={listview}>Badges</th>
                            <th style={listview}>Sent To</th>
                            <th style={listview}>Sent At</th>
                          </tr>
                        </thead>
                        <tbody>
                          <this.Sentitem />
                        </tbody>
                      </Table>
                    </div>
                  ) : ""
                }
  
              </div>
            </div>
        </Provider>
      );
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

