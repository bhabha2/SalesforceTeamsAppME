const axios = require("axios");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const ACData = require("adaptivecards-templating");
const COMMAND_ID = "getAccountOpportunities";
const ACAccountOpportunities = require("../adaptiveCards/ACAccountOpportunities.json");
  // Message extension Code
  // define function to search incident

  async function handleTeamsMessagingExtensionQuery(context, query,accessToken) {
    // Add your code here
    //await authorizeUser.authorizeUser();

    const searchQuery = query.parameters[0].value;
    let searchValue='';
    // look for 'incident_no', 'short_description' and 'assigned_to' in query and assign the value to SearchParameter and SearchValue
    // let readQuery = 'https://microsoftvalidation.my.salesforce.com.com/services/data/v60.0/parameterizedSearch/?q='+searchQuery+'&sobject=Account&Account.fields=id,name&Account.limit=10';
    // let readQuery = ' https://microsoftvalidation.my.salesforce.com/services/data/v60.0/parameterizedSearch/?q='+searchQuery;
    let readQuery = 'https://microsoftvalidation.my.salesforce.com/services/data/v59.0/query?q=SELECT+id,Name,(SELECT+Id,Name,CloseDate,StageName,Amount+FROM+Opportunities)+FROM+Account+WHERE+Name+LIKE+%27%25'+searchQuery+'%25%27';
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
          let opportunities = [];
          for (let i = 0; i < json.length; i++) {
            let item = json[i];
            for (let j = 0; j < item.Opportunities.records.length; j++) {
              let item2 = item.Opportunities.records[j];
              opportunities.push({
                  opportunityName:item2.Name,
                  opportunityCloseDate: item2.CloseDate,
                  opportunityStageName: item2.StageName,
                  opportunityAmount: item2.Amount,
              });
            }
            const template = new ACData.Template(ACAccountOpportunities);
            const resultCard = template.expand({
              $root: {
                accountName:item.Name,
                accountType: item.attributes.type,
                opportunities: opportunities,
              },
              });
            // console.log('\r\nresultCard: ',resultCard);
            const preview = CardFactory.heroCard(item.Name, item.attributes.type);
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
