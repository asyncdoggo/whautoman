If $CmdLine[0] Then ;Checks if there are any parameters provided
   $para = ""
   For $i = 1 To UBound($CmdLine)-1
       $para = $para & StringFormat(" %s%s%s", Chr(34),$CmdLine[$i], Chr(34)) ; Loop through parameters and add quotes before and after
   Next
   ControlFocus("Open","","Edit1");Brings the window in focus
   ControlSetText("Open","","Edit1",$para);Types the content present in para in the box
   ControlClick("Open","","Button1");Presses the button
EndIf