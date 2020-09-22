Attribute VB_Name = "modDictCompress"
' Dictionary Compression algorythm for compression of short
' English strings.  Designed for chat server, game server application
' 2008 - Timothy Patrick, timbo_m45@hotmail.com
' ===============================================================
' :::::: NOTES AND DIRECTIONS :::::::::::::::::::::::::::::::::::
' STEP 1) Call the function mdc_loaddicts to load dictionary data
' and initialize the compression engine
' STEP 2) Pass any string to the mdc_CompressText function
' and a compressed string will be returned

' To decrompress, pass a compressed string to mdc_DecompressText
' limitations:
' the five delimiter characters, and 3 flags as shown in the constants
' below (characters 248-255) cannot be compressed, and will cause
' a miss-interpretation to occur but no program errors should occur
' even if those characters are inserted.
' NEW FEATURE: Now upper/lower case are supported, however
' only 2 patterns are compressed, all small letters, ALL CAPITAL
' words and Words That Start With The First Letter As A Capital.
' Other words that do not fit these patterns are still usable, just
' not compressed. (for example: ThIs TyPE oF TeXt)
' The dictionaries can hold 247 words each,
' making for 741 words in total.
' The first dictionary is completely full*
'
'It is suggested you check that a word is not yet in the dictionary
'before adding a new word.  Do this by simply running the program
'and typing the word and seeing if it is compressed or if it stays
'as readable text within the compressed string; if it's readable, it
'is NOT in the dictionary.

'NOTE: This program looks at the beginnings of words so
'in some cases the words in the dictionary have been optimized to
'account for common word beginnings (eg. "possib" can be used to compress
'all of the following words: "possibility", "possible", "possibly", "possibilities")

' It's intended designation as a chat-room client-side compression
' is filled very nicely with the code in its current state.
'=================================================================

'Used for Compression mostly
Public mdcDICT() As String
Public mdcINDEX() As Long
Public mdcDICTPOS() As Long

'(used for everything, but mostly used during decompression)
Public Dict1StrArray() As String 'original dict array 1
Public Dict2StrArray() As String 'original dict array 2
Public Dict3StrArray() As String 'original dict array 3
Public Dict4StrArray() As String 'original dict array 4
Public Dict5StrArray() As String 'original dict array 5

'delimiters/reserved characters


Public CaseFlag1 As String  '= "ø"" 'chr(248)"
Public CaseFlag2 As String  '= "ù"" 'chr(249)"
Public CaseFlag3 As String  '= "ú"" 'chr(250)"

Public mdc_delim1 As String  '= "û"" 'chr(250) // dict1 delim"
Public mdc_delim2 As String  '= "ü"" 'chr(251) // dict2 delim"
Public mdc_delim3 As String  '= "ý" 'chr(252) // dict3 delim
Public mdc_delim4 As String  '= " // dict4 delim
Public mdc_delim5 As String  '= "þ" 'chr(253) // delim for neutral string'not in neutral mode
Public mdc_delim6 As String  '= "ÿ" 'chr(254) // delim for appending plain text to the end of current word
Public mdc_delim7 As String  '=

Public Sub mdc_loadDicts(Dict1FileNAMEPATH As String, Dict2FileNAMEPATH As String, Dict3FileNAMEPATH As String, Dict4FileNAMEPATH As String, Dict5FileNAMEPATH As String)
Dim intFile As Long  '<- stores freefile number
Dim strFile As String '<- stores dictionary before split
Dim i As Long '<- used in FOR..NEXT statements...
Dim j As Long '<- also used in FOR..NEXT statements..
Dim d As Long '<- used for tracking how many words we have in the sorted array
Dim SortLsize As Long 'biggest word's length...

intFile = FreeFile 'select free file number


'set delimiter characters
CaseFlag1 = Chr(0)
CaseFlag2 = Chr(1)
CaseFlag3 = Chr(2)

mdc_delim1 = Chr(3) '// dict1 delim
mdc_delim2 = Chr(4) '// dict2 delim"
mdc_delim3 = Chr(5) '// dict3 delim
mdc_delim4 = Chr(6) '// dict4 delim
mdc_delim5 = Chr(7) '// dict5 delim
mdc_delim6 = Chr(8) '// delim for neutral string'not in neutral mode
mdc_delim7 = Chr(9) '// delim for appending plain text to the end of current word

'load the dictionary data file into our temp string
'so we can procede to split it...
Open Dict1FileNAMEPATH For Input As #intFile
   strFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
Close #intFile

Dict1StrArray = Split(strFile, " ") 'Dict1StrArray just became an array of strings


'now load dict2
Open Dict2FileNAMEPATH For Input As #intFile
   strFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
Close #intFile

Dict2StrArray = Split(strFile, " ") 'Dict2StrArray just became an array of strings


'now load dict3
Open Dict3FileNAMEPATH For Input As #intFile
   strFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
Close #intFile

Dict3StrArray = Split(strFile, " ") 'Dict3StrArray just became an array of strings

'now load dict4
Open Dict4FileNAMEPATH For Input As #intFile
   strFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
Close #intFile

Dict4StrArray = Split(strFile, " ") 'Dict3StrArray just became an array of strings


'now load dict5
Open Dict5FileNAMEPATH For Input As #intFile
   strFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
Close #intFile

Dict5StrArray = Split(strFile, " ") 'Dict3StrArray just became an array of strings


'prepare to sort from largest to smallest word
'this is neccissary to make the dictionary lookup more
'effecient and so we can check the beginnings of words instead
'of just entire words

'STEP 1 = find the largest word size
For i = 0 To UBound(Dict1StrArray)
If Len(Dict1StrArray(i)) > SortLsize Then SortLsize = Len(Dict1StrArray(i))
Next
'.. also check the second dict
For i = 0 To UBound(Dict2StrArray)
If Len(Dict2StrArray(i)) > SortLsize Then SortLsize = Len(Dict2StrArray(i))
Next
'.. also check the third dict
For i = 0 To UBound(Dict3StrArray)
If Len(Dict3StrArray(i)) > SortLsize Then SortLsize = Len(Dict3StrArray(i))
Next
'.. also check the fourth dict
For i = 0 To UBound(Dict4StrArray)
If Len(Dict4StrArray(i)) > SortLsize Then SortLsize = Len(Dict4StrArray(i))
Next

'.. and also check the fourth dict
For i = 0 To UBound(Dict5StrArray)
If Len(Dict5StrArray(i)) > SortLsize Then SortLsize = Len(Dict5StrArray(i))
Next

'STEP 2 = go backwards from largest word to smallest
'and put into new array

'prepare dest. arrays
ReDim mdcDICT(0 To (UBound(Dict1StrArray) + UBound(Dict2StrArray) + UBound(Dict3StrArray) + UBound(Dict4StrArray) + UBound(Dict5StrArray) + 4)) As String
ReDim mdcINDEX(0 To UBound(mdcDICT)) As Long
ReDim mdcDICTPOS(0 To UBound(mdcDICT)) As Long
'
'Dim strTemper1 As String
'j = 0
'For i = 119 To UBound(Dict2StrArray)
'  strTemper1 = strTemper1 & Dict2StrArray(i) & " "
'  j = j + 1
'Next
'
'Clipboard.Clear
'Clipboard.SetText strTemper1
'MsgBox j
'Exit Sub

'//for every word size
For i = SortLsize To 1 Step -1

'// for every word in dict 1...
'// search for words of current sort-size and add to final array
  For j = 0 To UBound(Dict1StrArray)
    If Len(Dict1StrArray(j)) = i Then 'found a word that matches current size we're checking for
    mdcDICT(d) = Dict1StrArray(j)
    mdcINDEX(d) = 1
    mdcDICTPOS(d) = j  'index of word inside that dict...
    d = d + 1
    End If
  Next
  
  'rinse and repeat for dict2 data...
  For j = 0 To UBound(Dict2StrArray)
    If Len(Dict2StrArray(j)) = i Then 'found a word that matches current size we're checking for
    mdcDICT(d) = Dict2StrArray(j)
    mdcINDEX(d) = 2
    mdcDICTPOS(d) = j 'index of word inside that dict...
    d = d + 1
    End If
  Next
  
  'rinse and repeat for dict3 data...
  For j = 0 To UBound(Dict3StrArray)
    If Len(Dict3StrArray(j)) = i Then 'found a word that matches current size we're checking for
    mdcDICT(d) = Dict3StrArray(j)
    mdcINDEX(d) = 3
    mdcDICTPOS(d) = j 'index of word inside that dict...
    d = d + 1
    End If
  Next
  
  
  'rinse and repeat for dict4 data...
  For j = 0 To UBound(Dict4StrArray)
    If Len(Dict4StrArray(j)) = i Then 'found a word that matches current size we're checking for
    mdcDICT(d) = Dict4StrArray(j)
    mdcINDEX(d) = 4
    mdcDICTPOS(d) = j 'index of word inside that dict...
    d = d + 1
    End If
  Next
  
  
  'rinse and repeat for dict5 data...
  For j = 0 To UBound(Dict5StrArray)
    If Len(Dict5StrArray(j)) = i Then 'found a word that matches current size we're checking for
    mdcDICT(d) = Dict5StrArray(j)
    mdcINDEX(d) = 5
    mdcDICTPOS(d) = j 'index of word inside that dict...
    d = d + 1
    End If
  Next
Next 'done sorting..

Debug.Print "TOTAL DICTIONARY WORDS: " & UBound(mdcDICT) + 1
Debug.Print "-----------------------------------------"
Debug.Print "DICT 1: " & UBound(Dict1StrArray) + 1
Debug.Print "DICT 2: " & UBound(Dict2StrArray) + 1
Debug.Print "DICT 3: " & UBound(Dict3StrArray) + 1
Debug.Print "DICT 4: " & UBound(Dict4StrArray) + 1
Debug.Print "DICT 5: " & UBound(Dict5StrArray) + 1

End Sub

Public Function mdc_CompressText(InputTXT As String)
Dim strW As String 'for storing "compressed" words
Dim strF As String 'for storing total string

Dim CDM As Long 'Current Dictionary Mode
        '0 = out of dictionaries
        '1 = in dictionary1
        '2 = in dictionary2
        '3 = in dictionary3
        '4 = in dictionary4
Dim CCM As Long 'current case mode
        '0 = all lower case (straight from dict
        '1 = first letter caps
        '2 = all caps
        '3 = Does not fit any mask, cant compress

Dim tCCM As Long 'temp ccm (for testing)
Dim CCMb As String
Dim i As Long '<- used in FOR..NEXT statements...
Dim j As Long '<- also used in FOR..NEXT statements...
Dim WordArray() As String
Dim ccmSpace As Boolean

'first split the uncompressed text into an array of words
WordArray = Split(RTrim(InputTXT), " ")

mdc_CompressText = InputTXT

    For i = 0 To UBound(WordArray) 'for every word in the array
    
    ccmSpace = False
    
        'CASE SECTION
        'before we continue to check dictionaries.. lets look at the words case
        'and compare to the current case flag..
        tCCM = GetCCM(WordArray(i))
        
        'if the case state of this word is different from previous/default
        If tCCM <> CCM Then
           CCM = tCCM 'set flag to new ccm
           'and put marker into final string
                    Select Case CCM
                    Case 0
                        CCMb = CaseFlag1
                        'strF = strF & CaseFlag1 'all lower case
                    Case 1
                        CCMb = CaseFlag2
                        'strF = strF & CaseFlag2 'first char upper
                    Case 2
                        CCMb = CaseFlag3
                        'strF = strF & CaseFlag3 'all upper
                    Case 3
                        strF = strF & mdc_delim6 'unmatched (so put uncompressable delim)
                        CDM = 0
                        ccmSpace = True
                    End Select
        End If
    
    'DICT SECTION
    'set word string buffer to null
    strW = vbNullString
      If CCM < 3 Then
            'first cycle through dict and check for the word...
            For j = 0 To UBound(mdcDICT)
                
                'NOTE: replaced the following
                'CPU critical line of code with
                'StrComp call which is faster by about 50%.
                
                'If Left(LCase(WordArray(i)), Len(mdcDICT(j))) = mdcDICT(j) Then 'found at least the beginning of the word in one of the dictionaries
                If StrComp(Left(WordArray(i), Len(mdcDICT(j))), mdcDICT(j), vbTextCompare) = 0 Then
                    'first switch to dictionary if needed put tag
                    If CDM <> mdcINDEX(j) Then
                        CDM = mdcINDEX(j)

                        Select Case mdcINDEX(j)
                        Case 1: strW = strW & mdc_delim1 'dict 1 tag
                        Case 2: strW = strW & mdc_delim2 'dict 2 tag
                        Case 3: strW = strW & mdc_delim3 'dict 3 tag
                        Case 4: strW = strW & mdc_delim4 'dict 4 tag
                        Case 5: strW = strW & mdc_delim5 'dict 5 tag
                        End Select
                    End If
                    
                    'next put char for current dict entry
                    strW = strW & Chr((mdcDICTPOS(j) + 10))
                    
                    'next check if that captured the whole word or just the beginning..
                    If LCase(WordArray(i)) <> mdcDICT(j) Then '.. captured only beginning of word
                     strW = strW & mdc_delim7 'delim for appending plain text to the end of current word
                     strW = strW & Right(WordArray(i), Len(WordArray(i)) - Len(mdcDICT(j)))
                     strW = strW & mdc_delim7 'close append tag..
                    End If
                    
                    Exit For
                    'Debug.Print "i = " & i & "  J = " & j & "  wordarray(i) = " & WordArray(i) & "  mdcdict(j) = " & mdcDICT(j) _
                    & "  CDM = " & CDM & "  StrW = " & strW
                    
                End If
              
            Next 'check next word in dict if needed..
            
            'ONLY if the word is in a dict then add the case flag if needed.
            If CDM > 0 Then
                strF = strF & CCMb
                CCMb = ""
            End If
    
        End If 'end if CCM < 3
    
    If Len(strW) > 0 Then 'found word in dict somewhere, add built word to final string
    strF = strF & strW '<- done processing word
    Else 'didn't find word...
        If CDM <> 0 Then
        strF = strF & mdc_delim6 'delim for neutral string'not in neutral mode
        CDM = 0 'switch compressor's dict flag to 0 (none-compressed word)
        strW = WordArray(i)  'word added as-is
        strF = strF & strW 'add neutral word (and tags if needed) to final array...
        Else
        strW = WordArray(i) 'word added as-is
                If i = 0 Or ccmSpace = True Then
                strF = strF & strW 'add neutral word (and tags if needed) to final array...
                Else
                strF = strF & " " & strW 'add neutral word (and tags if needed) to final array...
                End If
        End If
    End If
    
    Next 'on to next word...

'after all done, return final value as stored in strF
'lastly before we send off our text, make sure the last character is not a flag (unneeded)
If Right(strF, 1) = mdc_delim7 Then strF = Left(strF, Len(strF) - 1)

mdc_CompressText = strF

Erase WordArray
End Function

Public Function mdc_DecompressText(strInputC As String)
Dim strF As String
Dim i As Long
Dim CurChar As String
Dim lcdm As Long 'hold previous cdm state
Dim Spacer As String
'NOTES:
'mdc_delim1 = switch to dict1
'mdc_delim2 = switch to dict2
'mdc_delim4 = switch to non-dict word (not compressable)
'mdc_delim6 = appendment tag to add endings to compressed roots
Dim CDM As Long 'current dict mode..
'0=out of dict..
'1=in dict1
'2=in dict2
'3=in dict3
'4=in dict4
Dim CCM As Long
Dim xbuff As String 'string buffer

'for every char in compressed string...
For i = 1 To Len(strInputC)

CurChar = Mid(strInputC, i, 1)

        If CurChar = CaseFlag1 Then
        CCM = 0
        ElseIf CurChar = CaseFlag2 Then
        CCM = 1
        ElseIf CurChar = CaseFlag3 Then
        CCM = 2
        Else

                    If CurChar = mdc_delim1 Then 'switch to dict1 from here..
                        CDM = 1
                    ElseIf CurChar = mdc_delim2 Then 'switch to dict2 from here..
                        CDM = 2
                    ElseIf CurChar = mdc_delim3 Then 'switch to dict3 from here..
                        CDM = 3
                    ElseIf CurChar = mdc_delim4 Then
                        CDM = 4
                    ElseIf CurChar = mdc_delim5 Then
                        CDM = 5
                    ElseIf CurChar = mdc_delim6 Then
                        CDM = 0 'set dict mode to none
                        If Len(strF) > 0 Then strF = strF & " "
                    ElseIf CurChar = mdc_delim7 Then 'appendment..
                        If CDM <> 6 Then
                        lcdm = CDM
                        CDM = 6
                        Else
                        CDM = lcdm
                        End If
                    Else 'non-delimiter character
                        
                        If Len(strF) = 0 Then xbuff = "" Else xbuff = " "
                        
                        Select Case CDM
                        Case 1
                        strF = strF & xbuff & CCMWord(Dict1StrArray(Asc(CurChar) - 10), CCM)
                        Case 2
                        strF = strF & xbuff & CCMWord(Dict2StrArray(Asc(CurChar) - 10), CCM)
                        Case 3
                        strF = strF & xbuff & CCMWord(Dict3StrArray(Asc(CurChar) - 10), CCM)
                        Case 4
                        strF = strF & xbuff & CCMWord(Dict4StrArray(Asc(CurChar) - 10), CCM)
                        Case 5
                        strF = strF & xbuff & CCMWord(Dict5StrArray(Asc(CurChar) - 10), CCM)
                        Case 0
                        strF = strF & CurChar
                        Case 6
                        strF = strF & CurChar
                        End Select
                    End If
        End If
Next

mdc_DecompressText = strF

End Function

Private Function GetCCM(CsWord As String) As Long
Dim GetCCMres As Long
Dim strTestCCM As String

        '0 = all lower case (straight from dict
        '1 = first letter caps
        '2 = all caps
        '3 = Does not fit any mask, cant compress

'check for 0 (all lower case)
strTestCCM = LCase(CsWord) 'set lower case version

'if all lower case
If strTestCCM = CsWord Then
    GetCCM = 0
    Exit Function
End If

'test for all caps
strTestCCM = UCase(CsWord)
If strTestCCM = CsWord Then
    GetCCM = 2 'all caps
    Exit Function
End If

'test for first letter being capped
strTestCCM = CapFirstChar(CsWord)
If strTestCCM = CsWord Then
    GetCCM = 1
    Exit Function
End If

GetCCM = 3 'unknown
End Function

'example input:   "this"
'example output:  "This"
Private Function CapFirstChar(CFCWord As String) As String
Dim CFCstring As String
'take the first letter and cap it and then append the rest of the word (with small letters) onto the end..
CFCstring = UCase(Left(CFCWord, 1)) & LCase(Right(CFCWord, Len(CFCWord) - 1))
CapFirstChar = CFCstring
End Function

Private Function CCMWord(sccmW As String, wCCM As Long) As String
Dim s As String
    
    Select Case wCCM
    Case 0 'all small
    s = LCase(sccmW)
    Case 1 'first caps
    s = CapFirstChar(sccmW)
    Case 2 'all caps
    s = UCase(sccmW)
    Case 3 'not processed
    s = sccmW
    End Select
    
    CCMWord = s
End Function

'CALL TO UNLOAD ALL COMPRESSOR VARS WHEN PROGRAM TERMINATES
Public Sub UnloadCompressor()
    Erase mdcDICT
    Erase mdcINDEX
    Erase mdcDICTPOS

    Erase Dict1StrArray
    Erase Dict2StrArray
    Erase Dict3StrArray
    Erase Dict4StrArray
End Sub
