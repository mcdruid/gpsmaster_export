#include <Constants.au3>

;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Drew Webber (mcdruid.co.uk)
;
; Script Function:
;   Opens GPS Master, traverses the tree of dates, exports each entry as a gpx file.
;

; Run GPS Master
Run("C:\Program Files\GPS Master 2.0.14\GPS Master.exe")

; Wait for the app to become active. The classname is monitored instead of the window title
WinWaitActive("[CLASS:WATCH_09059]")

Local $years
Local $months
Local $curmonth
Local $items

; Click on the first leaf in the SysTree
; https://www.autoitscript.com/autoit3/docs/functions/ControlTreeView.htm

; clunky way to do this, but assume nobody has been using GPS watch for 10 years :)
For $year = 0 To 10
  If ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Exists", "#"&$year ) Then
   ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Expand", "#"&$year )
   $months = ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "GetItemCount", "#"&$year )
   For $month = 0 To ($months - 1)

		   Switch $month
		   Case 10
			  $curmonth = "02"
		   Case 11
			  $curmonth = "01"
		   Case Else
			  $curmonth = "#"&$month
		   EndSwitch


	  ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Expand", "#"&$year&"|"&$curmonth )
	  $items = ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "GetItemCount", "#"&$year&"|"&$curmonth )
	    For $i = 0 To ($items - 1)
		   ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Select", "#"&$year&"|"&$curmonth&"|#"&$i )
		   ;MsgBox($MB_SYSTEMMODAL, "Debug", "year: "&$year&" month: "&$month&" i: "&$i)
		   ;If ((($year > 0) And ($month > 9)) Or ($year > 1)) Then
			  ;MsgBox($MB_SYSTEMMODAL, "Debug", "year: "&$year&" month: "&$month&" curmonth: "&$curmonth&" i: "&$i)
		     SaveAsGpx()
		   ;EndIf
	    Next
   Next

;ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Expand", "#0" )
;ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Expand", "#0|#0" )
;ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Select", "#0|#0|#0" )
;SaveAsGpx()

  EndIf
Next

; Now wait for GPS Master to close before continuing
WinWaitClose("[CLASS:WATCH_09059]")

; Finished!
Exit

Func SaveAsGpx()
   ; alt + F then E should bring up the Export dialogue
   Sleep(1000)
   Send("!f")
   Sleep(250)
   Send("e")
   ; choose a different format than the default .tkl
   ; https://www.autoitscript.com/autoit3/docs/functions/ControlCommand.htm
   ; couldn't get these to work
   ;ControlClick ( "[CLASS:#32770]", "", "[CLASS:ComboBox; INSTANCE:2]" )
   ;ControlCommand ( "[CLASS:#32770]", "", "[CLASS:ComboBox; INSTANCE:2]", "SelectString", "GPS Exchange file (*.gpx)" )
   ;Sleep(1000)
   WinWaitActive("[CLASS:#32770]") ; this is the "Save As" dialog
   Send("{TAB}")
   Send("{DOWN 4}") ; should select "GPS Exchange file (*.gpx)"
   Send("{ENTER}")
   Sleep(250)
   Send("!s") ; Save
EndFunc   ;==>SaveAsGpx
