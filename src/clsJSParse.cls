VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJSParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************************
'*                   RECURSIVE JSON TEXT PARSING CLASS                                    *
'*                                                                                        *
'* Authour: Mark Timieski                                                                 *
'* Date: 18OCT14                                                                          *
'*                                                                                        *
'* License: Code Project Open License (CPOL) 1.02                                         *
'*          http://www.codeproject.com                                                    *
'*                                                                                        *
'* Description:                                                                           *
'* Load JSON string through the 'Loadstring' method                                       *
'* Then access the parsed keys and values through the 'key' and 'Value' properties        *
'*                                                                                        *
'* Version:                                                                               *
'* 1.0 Initial Release                                                                    *
'*                                                                                        *
'******************************************************************************************

Private strKey As Variant
Private strVal As Variant
Private intHMax As Integer
Private lngStatus As Long

Private Sub Class_Initialize()
    lngStatus = -1
End Sub

Public Property Get err() As Long
'Status:
'   1 = JSON string has been sucessfully parsed
'  -1 = JSON string has not been loaded, no results are available
'  -2 = JSON string cannot be fully parsed (JSON text not fully or incorrectly formed)
    err = lngStatus
    
End Property

Public Property Get NumElements() As Integer
    NumElements = intHMax
End Property

Public Property Get Key(Index As Integer) As Variant
    If Index > UBound(strKey) Or Index < LBound(strKey) Then
        Key = ""
    Else
        Key = strKey(Index)
    End If
End Property

Public Property Get Value(Index As Integer) As Variant
    If Index > UBound(strVal) Or Index < LBound(strVal) Then
        Value = ""
    Else
        Value = strVal(Index)
    End If
End Property

Public Sub LoadString(JSONText As String)
'Load the JSON text into an array

    Const cLongMax = (2 ^ 31) - 1 'Maximum Value for Long type
    
    Dim lngIndex As Long
    Dim lngContLoc As Long
    Dim lngLoc As Long
    Dim lngDelimitOffset As Long
    Dim lngASize As Long
    
    Dim intNoOfChecks As Integer 'Number of different control characters in JSON
    Dim intCheck As Integer
    Dim intCtrlChr As Integer
    Dim intObJLvl As Integer
    Dim intAryElement As Integer
    Dim intLvl As Integer
    
    Dim strID As String
    Dim strChr As String
    Dim strKeyValue As String
    Dim strValue As String
    Dim strPChar As String
    Dim strFoundVal As String
    Dim strTempString As String
    Dim strAKey() As String
    Dim strAVal() As String
    Dim strALvlKey(100) As String
    
    Dim blArray As Boolean 'Flag to indicate that an array has been found
    Dim blValue As Boolean 'Falg to indicate that a value has been found
    Dim blKeyAndValue As Boolean 'Found a key and value pair
    Dim blDebug As Boolean
    
    'Set the flag to true if you want to see debug information
    'during the loading process
    blDebug = True
    
    On Error GoTo ErrHandler:
    
    lngASize = 10
    ReDim strAKey(lngASize)
    ReDim strAVal(lngASize)
    
    'intArrayElement = 1 'initialize value
    'initialize values
    blArray = False
    
    'Generate a string of control characters
    'String is {[:,]}"
    strID = ""
    strID = strID & Chr(123) 'The '{' character
    strID = strID & Chr(91)  'The '[' character
    strID = strID & Chr(58)  'The ':' character
    strID = strID & Chr(44)  'The ',' character
    strID = strID & Chr(93)  'The ']' character
    strID = strID & Chr(125) 'The '}' character
    strID = strID & Chr(34)  'The '"' character
    
    intNoOfChecks = Len(strID)
    intObJLvl = 0
    lngIndex = 1 'First element in the array will be strKey(1) and strVal(1)
    
    'As we process the JSON string it becomes shorter and shorter, until
    'its all been processed
    Do While Len(JSONText) > 0
        
        'Set to maximum value as default
        lngContLoc = cLongMax
        
        'Find Next control character:
        'Scan the text for the closest control character
        'to the beginning of the remaining JSON text
        For intCheck = 1 To intNoOfChecks
        
            strChr = Mid(strID, intCheck, 1)
            lngLoc = InStr(1, JSONText, strChr, vbBinaryCompare)
        
            If (lngLoc > 0) And (lngLoc < lngContLoc) Then
                lngContLoc = lngLoc
                intCtrlChr = intCheck
                strPChar = strChr
            End If
        
        Next intCheck
        
        'When the above for next loop ends we will have found the closest control character
        'stored in intCtrlChr - an index (1 to 8) to the found character in strChr
        'stored in lngContLoc - position of the next control character
        'stored in strPChar - the closest next control character
        
        If blDebug = True Then
            Debug.Print "Parse Character: " & strPChar
        End If
        
        'A control character has been found, figure out what to do by the found character
        If lngContLoc < cLongMax Then
         'Capture the information before the control character
         strValue = Mid(JSONText, 1, lngContLoc - 1)
         'Capture everything after the control character (the remaining JSON string)
         JSONText = Mid(JSONText, lngContLoc + 1, Len(JSONText))
        Else
            'We found the end of the JSON string
            Exit Do
        End If
        
        'Found an integer or boolean value or key (the comma)
        If (intCtrlChr = 4) Then
         If (blValue = True) Or (blArray = True) Then
            'Found a value, and we already have key
            strFoundVal = fnStringToVal(strValue)
            blKeyAndValue = True 'Set the "Key and value found" flag
          End If
        End If
        
        'Start of object (The "{" character)
        If intCtrlChr = 1 Then
            intObJLvl = intObJLvl + 1
            blArray = False 'An object, not an array
            blValue = False 'Need to find a key first
            If blDebug = True Then
                Debug.Print "Start of Object, Moved up to level" & intObJLvl
            End If
        End If
        
        'End of of object (The "}" character)
        If intCtrlChr = 6 Then
            intObJLvl = intObJLvl - 1
            blValue = False 'Need to find a key first
            If blDebug = True Then
                Debug.Print "End of Object, Moved down to level" & intObJLvl
            End If
        End If
        
        'Start of array (The "[" character)
        If intCtrlChr = 2 Then
            'intObJLvl = intObJLvl + 1
            'strALvlKey(intObJLvl) = intArrayElement
            blArray = True
            blValue = True 'Next thing should be a value
            intAryElement = 1
            If blDebug = True Then
                Debug.Print "Start of Array, Moved up to level" & intObJLvl
            End If
        End If
        
        'End of of array (The "]" character)
        If intCtrlChr = 5 Then
            'intObJLvl = intObJLvl - 1
            blArray = False
            blValue = False 'Need to find a key first
            If blDebug = True Then
                Debug.Print "End of Array, Moved down to level" & intObJLvl
            End If
        End If
        
        'Object Value start is found (The ":" character)
        If intCtrlChr = 3 Then
            blValue = True
            If blDebug = True Then
                Debug.Print "ready to get value"
            End If
        End If
        
        'Start of a string (the quote " character)
        'Can be a key or value
        If intCtrlChr = 7 Then
        
            'The start of the key or value has been found
            'The next quote will end the key or value
            '(unless the quote has an escape character in front of it "\")
            
            lngDelimitOffset = 1
          
            Do
                'Look for the next quote character
                lngLoc = InStr(lngDelimitOffset, JSONText, Chr(34), vbBinaryCompare)
                
                'If the string is zero length "" then exit the loop
                If lngLoc = 1 Then
                    Exit Do
                End If
            
                'Check to see if there is a delimter just before the quote
                'if there is then quote is part of the string and not the end of
                'the string.
                If Mid(JSONText, lngLoc - 1, 1) = Chr(92) Then
                    ' The quote character has an escape character in front of it
                    'so this quote doesn't count.  Remove the escape character.
                    JSONText = Mid(JSONText, 1, lngLoc - 2) & Mid(JSONText, lngLoc, Len(JSONText))
                    'and move the start of the check past the delimted quote
                    lngDelimitOffset = lngLoc
                    
                    'If we have a boogered JSON string where there is no valid closing quotes
                    'the above "if" will cause an error (the MID statement will attempt to check
                    'the string starting at a position of -1) and the code will jump to the error
                    'handling section.  If this error didn't occur the do..loop would get stuck.
    
                Else
                    Exit Do
                End If
            Loop
            
            'We now have a string, find any other delimiters
            '(any delimited " characters have already been fixed)
            strTempString = fnStringFix(Mid(JSONText, 1, lngLoc - 1))
            
            If (blValue = True) Or (blArray = True) Then
                'The key has been previously found and this is the value for the key
                strFoundVal = strTempString
                blKeyAndValue = True 'Set the "Key and value found" flag
            Else
                If lngLoc > 0 Then
                    'We've found a key
                    strALvlKey(intObJLvl) = strTempString
                    If blDebug = True Then
                        Debug.Print "Found Key:" & strALvlKey(intObJLvl) & _
                                    " for Level: " & intObJLvl
                    End If
                End If
            End If
            JSONText = Mid(JSONText, lngLoc + 1, Len(JSONText))
        End If
        
        
        'Found a key and value, move it to the array
        If blKeyAndValue = True Then
        
            If lngIndex > lngASize Then
                lngASize = lngASize + 100
                ReDim Preserve strAKey(lngASize)
                ReDim Preserve strAVal(lngASize)
            End If
        
            strAKey(lngIndex) = ""
            For intLvl = 1 To intObJLvl
                strAKey(lngIndex) = strAKey(lngIndex) & ">" & strALvlKey(intLvl)
            Next intLvl
            
            If (blArray = True) Then
                'add the array element to the key
                strAKey(lngIndex) = strAKey(lngIndex) & ">" & Trim(str(intAryElement))
                'increment the array element
                intAryElement = intAryElement + 1
            End If
            
            strAVal(lngIndex) = strFoundVal
            If blDebug = True Then
                Debug.Print "Added Key:" & strAKey(lngIndex) & " Value: " & strAVal(lngIndex)
            End If
            lngIndex = lngIndex + 1 'Increment the array
            blKeyAndValue = False 'Reset the "found" flag
            blValue = False 'Reset the "Value Found" flag
        End If
    
    Loop
    
    'Number of items found
    intHMax = lngIndex - 1
    strKey = strAKey
    strVal = strAVal
    lngStatus = 1 'JSON sucessfully parsed
Exit Sub
ErrHandler:
    
    'Error handling code
    lngStatus = -2 'JSON Parse error
    'Uncomment the next line to figure out the cause of the issue
    'Debug.Print VBA.err.Number
    'Debug.Print VBA.err.Description
    'Resume
    
End Sub


Private Function fnStringToVal(strInStr As String) As String
'Converts a string that contains formatting information into a string that only
'contains a value.  Values can be text, integer, or floating point values.
'null is pssed back as a zero length string: "".

    Dim intStrPos As Integer
    Dim strTemp As String
    Dim intChar As Integer
    
    'default value
    strTemp = ""
    
    'Make sure that the string does not have a zero length
    strInStr = " " & strInStr
    
    'Loop through each character in the string and remove anything
    'that is not alphanumeric.
    For intStrPos = 1 To Len(strInStr)
        intChar = asc(Mid(strInStr, intStrPos, 1))
        
        If ((intChar >= asc("a")) And (intChar <= asc("z"))) Or _
           ((intChar >= asc("A")) And (intChar <= asc("Z"))) Or _
           ((intChar >= asc("0")) And (intChar <= asc("9"))) Or _
           (intChar = asc(".")) Or (intChar = asc("+")) Or (intChar = asc("-")) Then
           
           strTemp = strTemp & Chr(intChar)
        End If
    
    Next intStrPos
    
    'Values that are listed as 'null' are converted to a zero length string
    If InStr(1, "null", strTemp, vbTextCompare) > 0 Then
        strTemp = ""
    End If
    
    fnStringToVal = strTemp

End Function

Private Function fnStringFix(strInput As String) As String
'This function goes through a JSON string and corrects delimited characters

Dim blParseComplete As Boolean
Dim lngStartPos As Long
Dim lngCurrentPos As Long

blParseComplete = False
lngStartPos = 1

Do While blParseComplete = False
    blParseComplete = True 'If we don't find any escape sequences then allo the loop to end
    
    'Escaped sequence: replace \\ with \
    lngCurrentPos = InStr(lngStartPos, strInput, "\\", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & "\" & _
                    Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

    'Escaped sequence: replace \/ with /
    lngCurrentPos = InStr(lngStartPos, strInput, "\/", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & "/" & _
                    Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

    'Escaped sequence: replace \b with a backspace
    lngCurrentPos = InStr(lngStartPos, strInput, "\b", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & Chr(8) & _
                    Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If
    
    'Escaped sequence: replace \f with a formfeed
    lngCurrentPos = InStr(lngStartPos, strInput, "\f", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & Chr(12) & _
                     Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

    'Escaped sequence: replace \n with a newline
    lngCurrentPos = InStr(lngStartPos, strInput, "\n", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & Chr(10) & _
                    Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

    'Escaped sequence: replace \r with a carriage return
    lngCurrentPos = InStr(lngStartPos, strInput, "\r", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & Chr(13) & _
                    Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

    'Escaped sequence: replace \t with a horizontal tab
    lngCurrentPos = InStr(lngStartPos, strInput, "\t", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & Chr(9) & _
                    Mid(strInput, lngCurrentPos + 2, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

    'Escaped sequence: replace \uXXXX with a unicode character
    lngCurrentPos = InStr(lngStartPos, strInput, "\u", vbTextCompare) 'look for the the specific escape sequence
    If lngCurrentPos > 0 Then
        strInput = Mid(strInput, 1, lngCurrentPos - 1) & _
                    ChrW$(CLng("&h" & Mid(strInput, lngCurrentPos + 2, 4))) & _
                    Mid(strInput, lngCurrentPos + 6, Len(strInput))
        blParseComplete = False 'set the status to check for another escape
    End If

Loop


fnStringFix = strInput
End Function


