Dim cam
Set cam = CreateObject("MaxIm.CCDCamera") 'create the application's CCDCamera object
Dim app
Set app = CreateObject("MaxIm.Application") ' Create a MaxIm DL Application Object


cam.LinkEnabled = True 'connect the camera
if Not cam.LinkEnabled Then 'if the connection failed raise a message
   wscript.echo "Failed to start camera."
   Quit
End If

cam.CoolerOn = True

app.FocuserConnected = True 'connect the focuser
if Not app.FocuserConnected Then 'if the connection failed raise a message
   wscript.echo "Failed to connect the focuser."
   Quit
End If

wscript.echo "Autofocus routine started."
app.AutoFocus()


Do While Not AutoFocusStatus
WScript.Sleep 1000
Loop
WScript.Echo "done"