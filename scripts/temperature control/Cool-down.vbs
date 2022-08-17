Dim cam 
Dim ho

Const InitStabTemp = 5	'Initial temprature stabilization in dg Celcius
Const LowTemp = -20	'Target temprature
Const PollTime = 3	'Interval for polling the temprature (seconds)
Const StepTime = 20	'Interval for cool-down (seconds)
Const StepTemp = 1	'Interval (dg Celcius)

Set cam = CreateObject("MaxIm.CCDCamera")
cam.LinkEnabled = True

if Not cam.LinkEnabled Then
   MsgBox "Failed to start camera."
   wscript.Quit
End If

cam.DisableAutoShutdown = True

MsgBox  "Camera is ready! Initiating temprature stabilization."

Set ho = CreateObject("DriverHelper.Util")
cam.CoolerOn = True
cam.TemperatureSetpoint = 50

'Do an initial temp stabilization (achieve a max 2-degree difference)
Do While Abs(cam.Temperature - InitStabTemp) > 2
	

	if (cam.CameraStatus = 1) or (cam.CoolerOn = False) Then
   		MsgBox "X Camera communication/cooler problem!"
   		wscript.Quit
	End If

	ho.WaitForMilliseconds(PollTime*1000)
Loop

cam.TemperatureSetpoint = InitStabTemp - StepTemp

MsgBox  "Initial temprature stabilization completed, proceeding to cool-down sequence"

Do While (cam.Temperature > LowTemp )
	ho.WaitForMilliseconds(StepTime*1000)

	if (cam.CameraStatus = 1) or (cam.CoolerOn = False) Then
   		MsgBox "X Camera communication/cooler problem!"
   		wscript.Quit
	End If

	'check if temprature is stabilized, then decrease, else wait another StepTime seconds
	if (cam.Temperature - cam.TemperatureSetpoint) <= 1 then
		cam.TemperatureSetpoint = cam.TemperatureSetpoint - StepTemp
	end if
	'wscript.echo "Camera temprature : " & cam.Temperature
Loop

cam.TemperatureSetpoint = LowTemp

MsgBox "Target temprature threshold reached! CCD Temprature reads : " & cam.Temperature




