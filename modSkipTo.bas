Attribute VB_Name = "modSkipTo"
Public Function skipTo(fileNum As Integer, str As String)
'This function searches a sequential text file with
'number fileNum until it finds str.  If no match is found,
'the function informs the user.
Dim buffer As String    'string buffer

Do
    Input #fileNum, buffer  'input from file
Loop Until buffer = str Or EOF(fileNum) 'loop until match found
                                        'or end of file reached

'If no match found, inform user
If EOF(fileNum) Then MsgBox str & " was not found in file #" & fileNum

End Function
