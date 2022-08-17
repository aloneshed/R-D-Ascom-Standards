Dim cam 
Dim ho

Const StepTime = 20	'Interval for warm-up (seconds)
Const StepTemp = 1	'Interval (dg Celcius)
Const AmbientTemp = 10	'Target temprature

Set cam = CreateObject("MaxIm.CCDCamera")
cam.LinkEnabled = True

if Not cam.LinkEnabled Then
   MsgBox "Failed to start camera."
   wscript.Quit
End If

cam.DisableAutoShutdown = True
cam.CoolerOn = True
MsgBox  "Camera is ready! Initiating warm-up sequence."
Set ho = CreateObject("DriverHelper.Util")

Do While (cam.Temperature < AmbientTemp)
	ho.WaitForMilliseconds(StepTime*1000)

	if (cam.CameraStatus = 1) or (cam.CoolerOn = False) Then
   		MsgBox "X Camera communication/cooler problem!"
   		wscript.Quit
	End If

	'check if temprature is stabilized, then increase, else wait another StepTime seconds
	if (cam.TemperatureSetpoint - cam.Temperature) <= 1 then
		cam.TemperatureSetpoint = cam.Temperature + StepTemp
	end if
Loop

cam.TemperatureSetpoint = AmbientTemp
MsgBox "Target temprature threshold reached! CCD Temprature reads : " & cam.Temperature


