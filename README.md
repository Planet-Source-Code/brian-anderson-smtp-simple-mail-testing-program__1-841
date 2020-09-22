<div align="center">

## SMTP: Simple Mail Testing Program


</div>

### Description

Allows sending of e-mail (SMTP) directly from a VB app using Winsock, WITH OUT having to buy an expensive add on componet
 
### More Info
 
Requires: Server Address (Name or IP), Senders & Recipeient's Names, Sender & Recipient E-Mail address, Body of message

Very straight forward. Makes sending mail from a VB program EASY!

Nothing really, does give status on sending operation

NONE!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Anderson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-anderson.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-anderson-smtp-simple-mail-testing-program__1-841/archive/master.zip)

### API Declarations

None


### Source Code

```
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim Start As Single, Tmr As Single
Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
 Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail per program start
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
 DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
 first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
 Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
 Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
 Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
 Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
 Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
 Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
 Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
 Eighth = Fourth + Third + Ninth + Fifth + Sixth ' Combine for proper SMTP sending
 Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
 Winsock1.RemoteHost = MailServerName ' Set the server address
 Winsock1.RemotePort = 25 ' Set the SMTP Port
 Winsock1.Connect ' Start connection
 WaitFor ("220")
 StatusTxt.Caption = "Connecting...."
 StatusTxt.Refresh
 Winsock1.SendData ("HELO yourdomain.com" + vbCrLf)
 WaitFor ("250")
 StatusTxt.Caption = "Connected"
 StatusTxt.Refresh
 Winsock1.SendData (first)
 StatusTxt.Caption = "Sending Message"
 StatusTxt.Refresh
 WaitFor ("250")
 Winsock1.SendData (Second)
 WaitFor ("250")
 Winsock1.SendData ("data" + vbCrLf)
 WaitFor ("354")
 Winsock1.SendData (Eighth + vbCrLf)
 Winsock1.SendData (Seventh + vbCrLf)
 Winsock1.SendData ("." + vbCrLf)
 WaitFor ("250")
 Winsock1.SendData ("quit" + vbCrLf)
 StatusTxt.Caption = "Disconnecting"
 StatusTxt.Refresh
 WaitFor ("221")
 Winsock1.Close
Else
 MsgBox (Str(Winsock1.State))
End If
End Sub
Sub WaitFor(ResponseCode As String)
 Start = Timer ' Time event so won't get stuck in loop
 While Len(Response) = 0
  Tmr = Start - Timer
  DoEvents ' Let System keep checking for incoming response **IMPORTANT**
  If Tmr > 50 Then ' Time in seconds to wait
   MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
   Exit Sub
  End If
 Wend
 While Left(Response, 3) <> ResponseCode
  DoEvents
  If Tmr > 50 Then
   MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
   Exit Sub
  End If
 Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub
Private Sub Command1_Click()
 SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
 'MsgBox ("Mail Sent")
 StatusTxt.Caption = "Mail Sent"
 StatusTxt.Refresh
 Beep
 Close
End Sub
Private Sub Command2_Click()
 End
End Sub
Private Sub Form_Load()
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 Winsock1.GetData Response ' Check for incoming response *IMPORTANT*
End Sub
```

