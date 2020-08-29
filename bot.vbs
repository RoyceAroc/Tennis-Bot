set WshShell=WScript.CreateObject("WScript.Shell")
WshShell.run "chrome.exe"
WScript.sleep 5000
WshShell.sendkeys "https://www.reservemycourt.com/login2.asp"
WshShell.sendkeys "{ENTER}"
WScript.sleep 8000
For x = 1 To 16
	WshShell.sendkeys "{TAB}"
	WScript.sleep 300
Next


WshShell.sendkeys "Janko"
WScript.sleep 300
WshShell.sendkeys "{TAB}"
WshShell.sendkeys "Lense"
WScript.sleep 300
WshShell.sendkeys "{ENTER}"
WScript.sleep 3000
WshShell.run "chrome.exe"
WScript.sleep 3000
Dim x
x = 5
Do While x < 6
'For testing but later change Hour(Now()) to 0  
 If Hour(Now()) = 0  Then
	Exit Do
 End If
Loop
If DatePart("w",Now()) > 4  Then
	'book court for Thursday, Friday, and Saturday
		
	a = DateAdd("d",1, Now)
	WshShell.sendkeys "https://www.reservemycourt.com/reservation.asp?invalidsubend=0&dtpick=" & Month(a) & "+52F" & Day(a) & "+52F" & Year(a) & "&courtpick=45198&startpick=6+53A00+53A00+=PM&endpick=7+53A00+53A00+=PM&repeatfreq=Weeks&repeatlimit=1&notetxt=&leaguepick=&courtcount=2&rt=0&checkin=&dailyres=0&compdt=&compyr=&edit=no&c_id=&onlinecourtfee=0&onlinecourtdeposit=0&offlinecourtfee=0&offlinecourtdeposit=0&wallet=personal&pay=true&transactiontypeid=1&singlemulti=Single"
	WScript.sleep 15000
	WshShell.sendkeys "{ENTER}"
	WScript.sleep 4000
	WshShell.sendkeys "+{TAB}"
	WScript.sleep 1500
	WshShell.sendkeys "{ENTER}"
	' Optional EndScript: Try reserving court 2 if court 1 might have failed reservation timing
	WshShell.run "chrome.exe"
	WScript.sleep 3000	
	a = DateAdd("d",1, Now)
	WshShell.sendkeys "https://www.reservemycourt.com/reservation.asp?invalidsubend=0&dtpick=" & Month(a) & "+52F" & Day(a) & "+52F" & Year(a) & "&courtpick=45198&startpick=6+53A00+53A00+=PM&endpick=7+53A00+53A00+=PM&repeatfreq=Weeks&repeatlimit=1&notetxt=&leaguepick=&courtcount=2&rt=0&checkin=&dailyres=0&compdt=&compyr=&edit=no&c_id=&onlinecourtfee=0&onlinecourtdeposit=0&offlinecourtfee=0&offlinecourtdeposit=0&wallet=personal&pay=true&transactiontypeid=1&singlemulti=Single"
	WScript.sleep 15000
	WshShell.sendkeys "{ENTER}"
	WScript.sleep 4000
	For x = 1 To 8
		WshShell.sendkeys "+{TAB}"
		WScript.sleep 300
	Next
	WshShell.sendkeys "{DOWN}"
	WScript.sleep 300				
	WshShell.sendkeys "{ENTER}"
	WScript.sleep 300
	WshShell.sendkeys "{TAB}"
	' For x = 1 To 22
		' WshShell.sendkeys "{DOWN}"
		' WScript.sleep 100
	' Next
	For x = 1 To 6
		WshShell.sendkeys "{TAB}"
		WScript.sleep 100
	Next
	WshShell.sendkeys "{ENTER}"
	WScript.sleep 1000
End If
CreateObject("WScript.Shell").Run "C:\Tennis-Bot\shut.vbs" 
