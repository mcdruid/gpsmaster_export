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

; stop at a given item (this item itself will not be exported)
Local $stopAt
; to stop at this item, uncomment this variable and set it to correct value
;$stopAt = "2014/02/22_11:47:41"

; Run GPS Master
Run("C:\Program Files\GPS Master 2.0.14\GPS Master.exe")

; Wait for the app to become active. The classname is monitored instead of the window title
WinWaitActive("[CLASS:WATCH_09059]")
WinActivate("[CLASS:WATCH_09059]") ; ensure focus

; Local $firstYear
Local $years
Local $months
Local $curmonth
Local $items
Local $curItem

; Find the first leaf in the SysTree
; https://www.autoitscript.com/autoit3/docs/functions/ControlTreeView.htm

; this seems to successfully find the text value of the first top-level item (the year)
;If ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Exists", "#0" ) Then
;  $firstYear = ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "GetText", "#0" )
;  MsgBox($MB_SYSTEMMODAL, "Debug", "firstYear: "&$firstYear )
;EndIf

; clunky way to do this, but assume nobody has been using GPS watch for 10 years :)
For $year = 0 To 10
  If ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Exists", "#"&$year ) Then
   ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "Expand", "#"&$year )
   $months = ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "GetItemCount", "#"&$year )
   For $month = 0 To ($months - 1)

      ; not sure if this is a bug in AutoIt, or whether I'm just doing something wrong
      ;   but using the item references #10 and #11 to select the last 2 months in the tree
      ;   doesn't seem to work for some reason. So map these to their text values and
      ;   use those instead.
      ; this might fail if items 10 and 11 aren't actually 02 and 01; that might happen
      ;   if there are entries for all months except for 02 (so 10 would then be 01).
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

         ; this is how to skip all entries until a given year / month (n.b. counting from 0).
         ;MsgBox($MB_SYSTEMMODAL, "Debug", "year: "&$year&" month: "&$month&" i: "&$i)
         ;If ((($year > 0) And ($month > 9)) Or ($year > 1)) Then
           ;MsgBox($MB_SYSTEMMODAL, "Debug", "year: "&$year&" month: "&$month&" curmonth: "&$curmonth&" i: "&$i)
         ; remember to add the EndIf after SaveAsGpx() if you use this option.

         ; alternatively you could stop here at a given year / month / item.
         $curItem = ControlTreeView ( "[CLASS:WATCH_09059]", "", "[CLASS:SysTreeView32; INSTANCE:1]", "GetText", "#"&$year&"|"&$curmonth&"|#"&$i )
         If $curItem == $stopAt Then
           MsgBox($MB_SYSTEMMODAL, "Debug", "Stopping at "&$curItem )
           ExitLoop 3 ; exit all loops traversing the tree
         EndIf

		     SaveAsGpx()
	    Next
   Next
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
   WinActivate("[CLASS:#32770]") ; ensure focus
   Send("{TAB}")
   Send("{DOWN 4}") ; should select "GPS Exchange file (*.gpx)"
   Send("{ENTER}")
   Sleep(250)
   Send("!s") ; Save
EndFunc   ;==>SaveAsGpx
