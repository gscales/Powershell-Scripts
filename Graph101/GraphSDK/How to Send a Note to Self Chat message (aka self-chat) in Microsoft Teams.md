To Send a Note to Self Chat message (aka self-chat) in Microsoft Teams you can use the following endpoint

https://graph.microsoft.com/v1.0/me/chats/48:notes/messages

The number prefix 48: prefix has special meaning in Teams eg 19: is for Channel 28: is for bots 29 for users (No referance documentation easliy availble) 

To Send a note to yourself using the Microsoft Graph Powershell SDK first connect with the scope to allow you to send Chat messages

```
connect-mggraph -Scopes "ChatMessage.Send"
```

Draft a JSON body with the Message you want to send
```
$MessagetoSend = "{
     `"body`": {
         `"content`": `"Take out the bins tonight`"
     }
 }"
```
Then Send the Message with
```
New-MgChatMessage -ChatId 48:notes -BodyParameter $MessagetoSend
```
