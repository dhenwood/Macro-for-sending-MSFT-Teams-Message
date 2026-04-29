import xapi from 'xapi';
import {getValidAccessToken} from './MicrosoftManageTokens';

// Set the following values to reflect each item
let teamsMessage = "For your current meeting, people are using the NSD5-24-Crecendo Room."
let teamsBotId = "teamsBotUuid"

// the following values should not be changed (global values)
let token = ""
let allUsers = []
let selectedUserId = ""

// INITIALISE LIST OF ALL USERS IN TENANT
async function listAllUsers() {
  let url = "https://graph.microsoft.com/v1.0/users"
  token = await getValidAccessToken()
  let header = `Authorization: Bearer ${token}`

  try {
    const response = await xapi.Command.HttpClient.Get({
      AllowInsecureHTTPS: "True",
      Header: header,
      ResultBody: "PlainText",
      Url: url
    });

    let responseBody = response.Body
    let responseJson = JSON.parse(responseBody)
    let value = responseJson.value
    value.forEach((item) => {
      let id = item.id
      let displayName = item.displayName
      let collection = [id, displayName]
      allUsers.push(collection)
    })
  } catch (err) {
    console.error('Refresh failed:', err);
  }
}

function searchUser(searchString) {
  let id = ""
  let name = ""
  allUsers.forEach((user) => {
    if (user[1].includes(searchString)) {
      id = user[0]
      name = user[1]
    }
  })
  return [id, name]
}

// STEPS TO SENDING MESSAGE
async function getChatId(id) {
  let url = "https://graph.microsoft.com/v1.0/chats"
  let contentHeader = "Content-Type: application/json"
  let authHeader = `Authorization: Bearer ${token}`
  let header = [contentHeader, authHeader]

  let body = `
  {
  "chatType": "oneOnOne",
  "members": [
    {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": ["owner"],
      "user@odata.bind": "https://graph.microsoft.com/v1.0/users('${teamsBotId}')"
    },
    {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": ["owner"],
      "user@odata.bind": "https://graph.microsoft.com/v1.0/users('${id}')"
    }
  ]
  }`

  try {
    const response = await xapi.Command.HttpClient.Post({
        AllowInsecureHTTPS: "True",
        Header: header,
        ResultBody: "PlainText",
        Url: url
      },
      body);

    let responseBody = response.Body
    let data = JSON.parse(responseBody)
    if (!data.id) {
      throw new Error('Invalid data response');
    } else {
      let id = data.id
      postMessage(id)
    }
  } catch (err) {
    console.error('Refresh2 failed:', err);
  }
}

async function postMessage(id) {
  let url = `https://graph.microsoft.com/v1.0/chats/${id}/messages`
  let contentHeader = "Content-Type: application/json"
  let authHeader = `Authorization: Bearer ${token}`
  let header = [contentHeader, authHeader]

  let body = `{
  "body": {
    "content": "${teamsMessage}"
  }
  }`

  try {
    const response = await xapi.Command.HttpClient.Post({
        AllowInsecureHTTPS: "True",
        Header: header,
        ResultBody: "PlainText",
        Url: url
      },
      body);

    let responseCode = response.StatusCode
    if (responseCode == "201") {
      console.log("message sent")
    } else {
      console.log("problem sending message: " + responseCode)
    }
  } catch (err) {
    console.error('Refresh failed:', err);
  }
}

// HANDLE SEARCH QUERY
async function displaySearchField() {
  xapi.Command.UserInterface.Message.TextInput.Display({
    FeedbackId: "searchFieldFeedback",
    Placeholder: "John Doe",
    SubmitText: "Search",
    Text: "Enter a persons name (or subset)",
    Title: "Directory Search"
  });
}

function confirmMessage(name) {
  const [id, username] = searchUser(name) ?? [];
  selectedUserId = id
  xapi.Command.UserInterface.Message.Prompt.Display({
    FeedbackId: "confirmBtn",
    "Option.1": "Confirm",
    Title: `Send ${username}`,
    Text: `About to send a Teams message to ${username}`
  });
}

function init() {
  listAllUsers()

  xapi.event.on('UserInterface Extensions Panel Clicked', (event) => {
    if (event.PanelId == 'intialButtonTrigger') {
      displaySearchField();
    }
  })

  xapi.event.on('UserInterface Message TextInput Response', (event) => {
    if (event.FeedbackId === 'searchFieldFeedback') {
      const responseSearchUser = event.Text;
      confirmMessage(responseSearchUser)
    }
  })

  xapi.event.on('UserInterface Message Prompt Response', (event) => {
    if (event.FeedbackId == "confirmBtn" && event.OptionId == 1) {
      getChatId(selectedUserId)
    }
  })
}


init()
