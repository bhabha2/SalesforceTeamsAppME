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
    console.log('readQuery: ',readQuery);
    console.log('config.url: ',config.url);
    try
    {
          const response = await axios.request(config);
          // const response = await runQuery(readQuery);
          console.log('\r\nresponse.data: ',JSON.stringify(response.data));
          const attachments = [];
          let json = response.data;

          for (let i = 0; i < json.length; i++) {
            let item = json[i];
            // console.log('\r\n',item);
            const template = new ACData.Template(ACAccountOpportunities);
            const resultCard = template.expand({
              $root: {
                name:item.records.name,
                type: item.records.type,
                //factset : item.records.Opportunities.records,
                //type, name, CloseDate, StageName,Amount
                
              },
              });
            console.log('\r\nresultCard: ',resultCard);
            const preview = CardFactory.heroCard(item.name, item.industry);
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
// {"totalSize":1,
// "done":true,
// "records":[
//   {"attributes":
//   {"type":"Account", "url":"/services/data/v59.0/sobjects/Account/001B000001OnByPIAV"},
//   "Id":"001B000001OnByPIAV",
//   "Name":"Air Tahiti",
//   "Opportunities":
//   {"totalSize":4,"done":true,"records":[
//     {"attributes":{"type":"Opportunity","url":"/services/data/v59.0/sobjects/Opportunity/006B0000007uoEhIAI"},"Id":"006B0000007uoEhIAI","Name":"Acme - 1,200 Widgets","CloseDate":"2019-11-13","StageName":"Value Proposition"},
//     {"attributes":{"type":"Opportunity","url":"/services/data/v59.0/sobjects/Opportunity/006B0000007uoEiIAI"},"Id":"006B0000007uoEiIAI","Name":"Acme - 600 Widgets","CloseDate":"2020-01-09","StageName":"Needs Analysis"},
//     {"attributes":{"type":"Opportunity","url":"/services/data/v59.0/sobjects/Opportunity/006B0000007uoEjIAI"},"Id":"006B0000007uoEjIAI","Name":"Acme - 200 Widgets","CloseDate":"2020-03-13","StageName":"Prospecting"},
//     {"attributes":{"type":"Opportunity","url":"/services/data/v59.0/sobjects/Opportunity/006B00000082RQ3IAM"},"Id":"006B00000082RQ3IAM","Name":"Engineering Services 787-9","CloseDate":"2022-11-30","StageName":"Negotiation/Review"}
//   ]}}]}