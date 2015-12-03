' This test was created using HP ALM
Window("Active and Today").Activate
Window("Active and Today").WinToolbar("ToolbarWindow32").Press "14", micLeftBtn
Window("Active and Today").Window("Block Details").Activate
Window("Active and Today").Window("Block Details").WinEdit("Security ID:").Set "MSFT"
Window("Active and Today").Window("Block Details").WinComboBox("Asset Class:").Select "EQ"
Window("Active and Today").Window("Block Details").WinComboBox("Trans Type:").Select "B"
Window("Active and Today").Window("Block Details").WinEdit("Size:").Set "600000"
Window("Active and Today").Window("Block Details").WinEdit("Price:").Set "15"
Window("Active and Today").Window("Block Details").WinEdit("Commission:").Set "0.03"
Window("Active and Today").Window("Block Details").WinComboBox("Security Type:").Select "EQU"
Window("Active and Today").Window("Block Details").WinComboBox("Oasys Broker:").Select "Jefferies & Company (JEFCO)"

For Iterator = 1 To 600 Step 1
	num = threeDigit(Iterator)
	subaccount = "G3_Acct" & num
	
	Print subaccount
	Call addAllocation(subaccount)
Next

Window("Active and Today").Window("Block Details").WinButton("OK").Click

Function threeDigit(x)
	If len(x) = 1 Then
		x = "00" & x
		threeDigit = x
	End If
	
	If len(x) = 2 Then
		x = "0" & x
		threeDigit = x		
	End If
	
	If len(x) = 3 Then
		threeDigit = x
	End If
End Function

Sub addAllocation(subaccount)
	Window("Active and Today").Window("Block Details").WinButton("New").Click
	Window("Active and Today").Dialog("Allocation Details").Activate
	Window("Active and Today").Dialog("Allocation Details").WinComboBox("ComboBox").Select subaccount
	Window("Active and Today").Dialog("Allocation Details").WinEdit("Size:").Set "1000"
	Window("Active and Today").Dialog("Allocation Details").WinButton("OK").Click
End Sub
