Rem VBS code for D3 VX series backplane reset
 
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity",,48)
foundhub = false
For Each objItem in colItems
    if instr(objItem.DeviceID, "VID_0451&PID_2046") Then
      foundhub = true
    End If
Next
if foundhub = true Then
  msgbox("Backplane Found")
Else
  msgbox("Backplane NOT FOUND")
end if