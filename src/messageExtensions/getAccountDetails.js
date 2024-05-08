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
          try
          {
          const response = await axios.request(config);

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
