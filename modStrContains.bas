Attribute VB_Name = "modStrContains"
Public Function strContains1(str As String, word As String)
'Determines if the string word is contained within str
Dim i As Integer  'counter

strContains1 = False    'initial assumption

'Loop through each character of str, creating a string of
'the same length as word, and compare
For i = 1 To Len(str) - (Len(word) - 1)
    'Convert all strings to uppercase - case insensitive
    If UCase$(word) = UCase$(Mid$(str, i, Len(word))) Then
        strContains1 = True  'match found
        i = Len(str)         'exit loop
    End If
Next i

End Function

Public Function strContains2(str As String, words() As String, n As Integer) As Boolean
'Determines if any of the n strings in words() is contained
'within str
Dim i, j As Integer   'counters

strContains2 = False   'initial assumption

'For each of the n strings in words(), parse str
'and look for match
For j = 1 To n  'loop through words()
    For i = 1 To Len(str) - (Len(words(j)) - 1) 'parse string
        'Convert all strings to uppercase (case insensitive)
        If UCase$(words(j)) = UCase$(Mid$(str, i, Len(words(j)))) Then
            strContains2 = True   'If match is found,
            i = Len(str)         'exit loop
            j = n
        End If
    Next i
Next j

End Function

