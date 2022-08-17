'Stephen Mounioloux - Cooledpix.com
'Refocusing Script V1.0 for MaximDL imaging (autosave) sequences
'2012/10/17 - Last update: 2012/10/18
'----------------------------------------------------------
Dim cam
Set cam = CreateObject("MaxIm.CCDCamera")
Dim app
Set app = CreateObject("MaxIm.Application")

cam.LinkEnabled = True
if Not cam.LinkEnabled Then
   wscript.echo "Failed to start camera."
   Quit
End If


app.CreateCalibrationGroups("//", 2, 1, False)



app.FocuserConnected = True
if Not af.FocuserConnected Then
   wscript.echo "Failed to connect the focuser."
   Quit
End If

wscript.echo "Moving filter wheel to Focus on Slot #0 Filter."
cam.filter(0)
wscript.echo "Autofocus routine started."
af.AutoFocus()

Do While Not AutoFocusStatus
WScript.Sleep 1000
Loop
