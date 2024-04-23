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
            const template = new ACData.Template(ACOpportunitiesByOwner);
            const resultCard = template.expand({
              $root: {
                name:item.name,
                type: item.type,
                industry: item.industry || '',
                website: item.website,
                owner: item.Owner.name || '',
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
// {"totalSize":23,"done":true,"records":[{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQqIAM"},"Id":"006B00000082RQqIAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001PDHGQIA5"},"Name":"Adventure Works"},"Name":"100K Spring Order","CloseDate":"2022-04-29","StageName":"Negotiation/Review","Amount":250000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQrIAM"},"Id":"006B00000082RQrIAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001PDHHEIA5"},"Name":"Microsoft Corporation"},"Name":"50K Fall Widget Order","CloseDate":"2021-10-04","StageName":"Closed Won","Amount":100000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQHIA2"},"Id":"006B00000082RQHIA2","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001P06pUIAR"},"Name":"Tailspin Toys"},"Name":"10K Winter Order","CloseDate":"2022-01-07","StageName":"Closed Won","Amount":50000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RPxIAM"},"Id":"006B00000082RPxIAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001P06pUIAR"},"Name":"Tailspin Toys"},"Name":"100K Widgets","CloseDate":"2022-06-30","StageName":"Prospecting","Amount":250000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RR0IAM"},"Id":"006B00000082RR0IAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001PDHHEIA5"},"Name":"Microsoft Corporation"},"Name":"10K Summer True-Up","CloseDate":"2022-07-29","StageName":"Qualification","Amount":20000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007UVIuIAO"},"Id":"006B0000007UVIuIAO","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001P06pUIAR"},"Name":"Tailspin Toys"},"Name":"25K Widgets","CloseDate":"2022-06-30","StageName":"Qualification","Amount":30000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000008k2JGIAY"},"Id":"006B0000008k2JGIAY","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001SbJy8IAF"},"Name":"Buy N Large Corp"},"Name":"FY23 BNL | OnBase Oppty","CloseDate":"2023-01-31","StageName":"Needs Analysis","Amount":null},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQRIA2"},"Id":"006B00000082RQRIA2","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByRIAV"},"Name":"Global Media"},"Name":"200K Fall Order","CloseDate":"2021-11-01","StageName":"Closed Won","Amount":500000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQSIA2"},"Id":"006B00000082RQSIA2","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByQIAV"},"Name":"salesforce.com"},"Name":"100K Winter Order","CloseDate":"2022-02-01","StageName":"Closed Won","Amount":250000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQ2IAM"},"Id":"006B00000082RQ2IAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001P06pUIAR"},"Name":"Tailspin Toys"},"Name":"10K Gadgets Summer Order","CloseDate":"2022-06-30","StageName":"Prospecting","Amount":20000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQ3IAM"},"Id":"006B00000082RQ3IAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByPIAV"},"Name":"Air Tahiti"},"Name":"Engineering Services 787-9","CloseDate":"2022-11-30","StageName":"Negotiation/Review","Amount":20000000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQCIA2"},"Id":"006B00000082RQCIA2","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001P06pUIAR"},"Name":"Tailspin Toys"},"Name":"100K Gadgets","CloseDate":"2022-04-29","StageName":"Prospecting","Amount":200000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQlIAM"},"Id":"006B00000082RQlIAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001PDHGQIA5"},"Name":"Adventure Works"},"Name":"20K Fall Order","CloseDate":"2021-11-01","StageName":"Closed Won","Amount":50000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B00000082RQvIAM"},"Id":"006B00000082RQvIAM","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001PDHHEIA5"},"Name":"Microsoft Corporation"},"Name":"100K Gadget Upgrade","CloseDate":"2022-06-30","StageName":"Negotiation/Review","Amount":500000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000008P4PbIAK"},"Id":"006B0000008P4PbIAK","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001QR9T2IAL"},"Name":"Demo JP Corp"},"Name":"Supper Big Deal","CloseDate":"2022-09-02","StageName":"Negotiation/Review","Amount":100000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEeIAI"},"Id":"006B0000007uoEeIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByRIAV"},"Name":"Global Media"},"Name":"salesforce.com - 5000 Widgets","CloseDate":"2022-10-11","StageName":"Id. Decision Makers","Amount":500000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEfIAI"},"Id":"006B0000007uoEfIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByRIAV"},"Name":"Global Media"},"Name":"salesforce.com - 500 Widgets","CloseDate":"2019-10-11","StageName":"Closed Won","Amount":50000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEgIAI"},"Id":"006B0000007uoEgIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByRIAV"},"Name":"Global Media"},"Name":"Global Media - 400 Widgets","CloseDate":"2019-12-12","StageName":"Id. Decision Makers","Amount":40000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEhIAI"},"Id":"006B0000007uoEhIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByPIAV"},"Name":"Air Tahiti"},"Name":"Acme - 1,200 Widgets","CloseDate":"2019-11-13","StageName":"Value Proposition","Amount":60000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEiIAI"},"Id":"006B0000007uoEiIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByPIAV"},"Name":"Air Tahiti"},"Name":"Acme - 600 Widgets","CloseDate":"2020-01-09","StageName":"Needs Analysis","Amount":70000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEjIAI"},"Id":"006B0000007uoEjIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByPIAV"},"Name":"Air Tahiti"},"Name":"Acme - 200 Widgets","CloseDate":"2020-03-13","StageName":"Prospecting","Amount":20000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoEkIAI"},"Id":"006B0000007uoEkIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByQIAV"},"Name":"salesforce.com"},"Name":"salesforce.com - 1,000 Widgets","CloseDate":"2019-11-13","StageName":"Negotiation/Review","Amount":100000},{"attributes":{"type":"Opportunity","url":"/services/data/v60.0/sobjects/Opportunity/006B0000007uoElIAI"},"Id":"006B0000007uoElIAI","Account":{"attributes":{"type":"Account","url":"/services/data/v60.0/sobjects/Account/001B000001OnByQIAV"},"Name":"salesforce.com"},"Name":"salesforce.com - 2,000 Widgets","CloseDate":"2020-01-11","StageName":"Value Proposition","Amount":20000}]}