const axios = require("axios");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const ACAccountDetails = require("../adaptiveCards/ACAccountDetails.json");
const COMMAND_ID = "getAccountDetails";
// const { runQuery } = require('./CommonFunctions');
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
    console.log('\r\nInside getAccountDetails');
    // Add your code here
    //await authorizeUser.authorizeUser();

    const searchQuery = query.parameters[0].value;
    let searchValue='';
    let readQuery = 'https://microsoftvalidation.my.salesforce.com/services/data/v60.0/parameterizedSearch/?q='+searchQuery+'&sobject=Account&Account.fields=id,name,type,industry,website,Owner.Name&Account.limit=10';
        let config = {
      method: 'get',
      maxBodyLength: Infinity,
      url: readQuery,
      headers: { 
        'Authorization': `Bearer ${accessToken}`, 
        'Cookie': ''    
      },
    };
            // console.log('readQuery: ',readQuery);
            // console.log('config.url: ',config.url);
        // axios.request(config)
        // .then((response) => {
          try
          {
          const response = await axios.request(config);
          // const response = await runQuery(readQuery);
          // console.log('\r\nresponse.data: ',JSON.stringify(response.data));
          const attachments = [];
          let json = response.data.searchRecords;
          for (let i = 0; i < json.length; i++) {
            let item = json[i];
            // console.log('\r\n',item);
            const template = new ACData.Template(ACAccountDetails);

            const resultCard = template.expand({
              $root: {
                id:item.Id,
                name: item.Name,
                type: item.Type,
                industry: item.Industry || '',
                website: item.Website,
                owner: item.Owner.Name || '',
              },
              });
            // console.log('\r\nresultCard: ',resultCard);
            const preview = CardFactory.heroCard(item.Name, item.Type);
            const attachment = { ...CardFactory.adaptiveCard(resultCard), preview };
            attachments.push(attachment);
          }
        return {
          composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments: attachments,
          },
        };
  
      // })
      }
      catch(error) {
        console.log(error);
      }
}
module.exports ={ COMMAND_ID, handleTeamsMessagingExtensionQuery };


// response.data:  
// {"searchRecords":[
//   {"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByPIAV"},
//   "Id":"001B000001OnByPIAV",
//   "Name":"Air Tahiti",
//   "Type":"Prospect",
//   "Industry":"Manufacturing",
//   "Website":"www.acme.com",
//   "Owner":{"attributes":{"type":"User","url":"/services/data/v60.0/sobjects/User/005B0000008A3YtIAK"},"Name":"Cameron Davis"}
// }]}
// resultCard:  {type: 'AdaptiveCard','$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',version: '1.5',body: [{type: 'TextBlock',text: '${user}',wrap: true,style: 'heading',size: 'Large'},{ type: 'Container', items: [Array] },{type: 'TextBlock',text: '001B000001OnByPIAV',wrap: true,style: 'default',size: 'Default',isVisible: 'false'}]}
