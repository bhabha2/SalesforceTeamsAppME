const axios = require("axios");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const COMMAND_ID = "getOpportunitiesByOwner";
const ACOpportunitiesByOwner = require("../adaptiveCards/ACOpportunitiesByOwner.json");
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
    // Add your code here
    //await authorizeUser.authorizeUser();

    const searchQuery = query.parameters[0].value;
    let searchValue='';
    // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
    let readQuery = 'https://microsoftvalidation.my.salesforce.com/services/data/v60.0/query?q=SELECT+Id,Account.Name,Name,CloseDate,StageName,Amount+FROM+Opportunity+WHERE+Owner.Name+LIKE+%27%25'+searchQuery+'%25%27'
    // let readQuery = 'https://microsoftvalidation.my.salesforce.com/services/data/v60.0/parameterizedSearch/?q='+searchQuery+'&sobject=Account&Account.fields=id,name,type,industry,website,Owner.Name&Account.limit=10';
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
          let json = response.data.records;
          for (let i = 0; i < json.length; i++) {
            let item = json[i];
            const template = new ACData.Template(ACOpportunitiesByOwner);
            const resultCard = template.expand({
              $root: {
                accountName:item.Account.Name,
                opportunityName:item.Name,
                opportunityCloseDate:item.CloseDate || '',
                opportunityStageName: item.StageName || 'Not Available',
                opportunityAmount: item.Amount || 'Not Available',
              },
              });
            const preview = CardFactory.heroCard(item.Account.Name, item.Name);
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
