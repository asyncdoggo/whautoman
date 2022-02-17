Func ConvertString($str)
   $newStr = """"; Manupluated string to return

   For $i = 0 To StringLen($str);Iterate over passed string
	  $char = StringMid($str,$i,1);Get each character
	  If $char == Chr(32) Then;If there is whitespace
	  $newStr = $newStr & String(""" """);insert double quotes
   Else
	  $newStr = $newStr & $char;Else insert the character
   EndIf

   Next

$newStr = $newStr & String("""")

Return $newStr
EndFunc


Func ConvertString_X($str)
   $newStr = """"

   $splitStr = StringSplit($str," ")

   For $i = 1 To $splitStr[0]
	  $newStr = $newStr & $splitStr[$i]
	  $newStr = $newStr & String(""" """)
   Next
$newStr = $newStr & String("""")
Return $newStr
EndFunc

If $CmdLine[0] > 0 Then
   $argument = String($CmdLine[1]); Convert to string

   $stringLen = StringLen($argument); Get the length of string

   $newStr = ConvertString_X($argument)

   ControlFocus("Open","","Edit1")
   ControlSetText("Open","","Edit1",$newStr)
   ControlClick("Open","","Button1")

Endif





