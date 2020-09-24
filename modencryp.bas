Attribute VB_Name = "modencryp"
Option Explicit


Dim i As Integer

'WhiteKnight 2000
'6-1-00
'If you have any comments of sugestions please e-mail them _
 me at witenite87@excite.com.  Feel Free to use this in any way you want.  You can Modify it any way _
 you like as long as the comments remain and you give credit where credit is do. _
 Please Visit my site http://camalot.virtualave.net.
'If you have any thing you would like me to add to this just email it to me and If I add it your name And Info _
 Will Go under each Sub/Function you add.
Public Function Encrypt(ByVal Plain As String, sEncKey As String) As String
    'Coded WhiteKnight 6-1-00
    'This Encrypts A string by converting it to its ASCII number but the difference is it uses a Key String it converts the keystring
    'to ASCII and adds it to the first ASCII Value the key is needed to decrypt the text.  I do plain on changing this some what but For
    'Now its ok.  I've only seen it cause an error when the wrong Key was entered while decrypting.
    'Note That If you use the same letter more then 3 times in a row then each letter after it if still the same is ignored
    '(ie aaa = aaaaaaaaa but aaa <> aaaza)
    'If anyone Can figure out a way to fix this please e-mail me
    Dim encrypted2 As String
    Dim LenLetter As Integer
    Dim Letter As String
    Dim KeyNum As String
    Dim encstr As String
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long
    
    On Error GoTo oops
    
    If sEncKey = "" Then sEncKey = "WhiteKnight" 'Sets the Encryption Key if one is not set
    ReDim encKEY(1 To Len(sEncKey))
    
    
    For i% = 1 To Len(sEncKey$) 'starts the values for the Encryption Key
        KeyNum = Mid$(sEncKey$, i%, 1) 'gets the letter at index i%
        encKEY(i%) = Asc(KeyNum) 'sets the the Array value to ASC number for the letter
        If i% = 1 Then Math = encKEY(i%): GoTo nextone 'This is the fist letter so just hold the value
        If i% >= 2 And Math - encKEY(i%) >= 0 And encKEY(i%) <= encKEY(i% - 1) Then Math = Math - encKEY(i%) 'compairs the value to the previous value and then either adds/subtracts the value to the Math total
        If i% >= 2 And Math - encKEY(i%) >= 0 And encKEY(i%) <= encKEY(i% - 1) Then Math = Math - encKEY(i%)
        If i% >= 2 And encKEY(i%) >= Math And encKEY(i%) >= encKEY(i% - 1) Then Math = Math + encKEY(i%)
        If i% >= 2 And encKEY(i%) < Math And encKEY(i%) >= encKEY(i% - 1) Then Math = Math + encKEY(i%)
nextone:
    Next i%
    
    Plain$ = scramb(Plain$)
    frmMain.txt_stp1.Text = Plain$
    For i% = 1 To Len(Plain) 'Now for the String to be encrypted
        Letter = Mid$(Plain, i%, 1) 'sets Letter to the letter at index i%
        LenLetter = Asc(Letter) + Math 'Now it adds the Asc value of Letter to Math
        If LenLetter >= 100 Then encstr = encstr & Asc(Letter) + Math & " " 'checks and corrects the format then adds a space to separate them frm each other
        If LenLetter <= 99 Then encstr$ = encstr & "0" & Asc(Letter) + Math & " " 'checks and corrects the format then adds a space to separate them frm each other
    Next i%


    frmMain.txt_stp2.Text = encstr
    
    'This is part of what i%'m doing to convert the encrypted numbers to  Letters so it sort of encrypts the encrypted message.
    temp$ = encstr 'hold the encrypted data
    temp$ = TrimSpaces(temp) 'get rid of the spaces
    itempnum% = Mid(temp, 1, 2) 'grab the first 2 numbers
    temp2$ = Chr(itempnum% + 100) 'Now add 100 so it will be a valid char
    
    If Len(itempnum%) >= 2 Then itempstr$ = Str(itempnum%) 'If its a 2 digit number hold it and continue
    If Len(itempnum%) = 1 Then itempstr$ = "0" & TrimSpaces(Str(itempnum%)) 'If the number is a single digit then add a '0' to the front then hold it
    
    
    encrypted2$ = temp2 'set the encrypted message
    frmMain.txt_stp3.Text = Asc(temp2$) & " "
    
    For i% = 3 To Len(temp) Step 2
        itempnum% = Mid(temp, i%, 2) 'grab the next 2 numbers
        temp2$ = Chr(itempnum% + 100) ' add 100 so it will be a valid char
        If i% = Len(temp) Then itempstr$ = Str(itempnum%): GoTo itsdone 'if its the last number we only want to hold it we don't want to add a '0' even if its a single digit
        If Len(itempnum%) = 2 Then itempstr$ = Str(itempnum%) 'If its a 2 digit number hold it and continue
        If Len(TrimSpaces(Str(itempnum))) = 1 Then itempstr$ = "0" & TrimSpaces(Str(itempnum%)) 'If the number is a single digit then add a '0' to the front then hold it
        'Now check to see if a - number was created if so cause an error message
        If Left(TrimSpaces(Str(itempnum)), 1) = "-" Then MsgBox "there is a bug in the ecryption method please contact whiteknight @ witenite87@excite.com with the following information." & vbCrLf & "Your Encryption Key, The Unencrypted String and the Encrypted String.", vbApplicationModal + vbCritical + vbDefaultButton1, "Encryption Error"
itsdone:
        frmMain.txt_stp3.Text = frmMain.txt_stp3.Text & Asc(temp2$) & " "
        encrypted2$ = encrypted2 & temp2$ 'Set The Encrypted message
        
    Next i%


    'Encrypt = encstr 'Returns the First Encrypted String
    Encrypt = encrypted2 'Returns the Second Encrypted String
    frmMain.txt_stp4.Text = Encrypt$
    Exit Function 'We are outta Here
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
End Function

Public Function Decrypt(ByVal Encrypted As String, sEncKey As String) As String
    'Coded By WhiteKnight 6-1-00
    'This Encrypts A string by converting it to its ASCII number but the difference is it uses a Key String it converts the keystring
    'to ASCII and adds it to the first ASCII Value the key is needed to decrypt the text.  I do plain on changing this some what but For
    'Now its ok.  I've only seen it cause an error when the wrong Key was entered while decrypting.
    'Note That If you use the same letter more then 3 times in a row then each letter after it if still the same is ignored
    '(ie aaa = aaaaaaaaa but aaa <> aaazaaa)
    'If anyone Can figure out a way to fix this please e-mail me
    Dim NewEncrypted As String
    Dim Letter As String
    Dim KeyNum As String
    Dim EncNum As String
    Dim encbuffer As Long
    Dim strDecrypted As String
    Dim Kdecrypt As String
    Dim lastTemp As String
    Dim LenTemp As Integer
    Dim temp As String
    Dim temp2 As String
    Dim itempstr As String
    Dim itempnum As Integer
    Dim Math As Long

    On Error GoTo oops

    If sEncKey = "" Then sEncKey = "WhiteKnight"

    ReDim encKEY(1 To Len(sEncKey))
    
    'Convert The Key For Decryption
    For i% = 1 To Len(sEncKey$)
        KeyNum = Mid$(sEncKey$, i%, 1) 'Get Letter i%% in the Key
        encKEY(i%) = Asc(KeyNum) 'Convert Letter i% to Asc value
        If i% = 1 Then Math = encKEY(i%): GoTo nextone 'if it the first letter just hold it
        If i% >= 2 And Math - encKEY(i%) >= 0 And encKEY(i%) <= encKEY(i% - 1) Then Math = Math - encKEY(i%) 'compairs the value to the previous value and then either adds/subtracts the value to the Math total
        If i% >= 2 And Math - encKEY(i%) >= 0 And encKEY(i%) <= encKEY(i% - 1) Then Math = Math - encKEY(i%)
        If i% >= 2 And encKEY(i%) >= Math And encKEY(i%) >= encKEY(i% - 1) Then Math = Math + encKEY(i%)
        If i% >= 2 And encKEY(i%) < Math And encKEY(i%) >= encKEY(i% - 1) Then Math = Math + encKEY(i%)
nextone:
    Next i%
    
    
    'This is part of what i'm doing to convert the encrypted numbers to  Letters so it sort of encrypts the encrypted message.
    temp$ = Encrypted 'hold the encrypted data

    For i% = 1 To Len(temp)
        itempstr = TrimSpaces(Str(Asc(Mid(temp, i%, 1)) - 100)) 'grab the next 2 numbers
        If Len(itempstr$) = 2 Then itempstr$ = itempstr$ 'If its a 2 digit number hold it and continue
        If i% = Len(temp) - 2 Then LenTemp% = Len(Mid(temp2, Len(temp2) - 3))
        'If LenTemp% <> 4 And i% = Len(temp) Then itempstr$ = "0" & TrimSpaces(itempstr$): MsgBox "Added": GoTo itsdone
        If i% = Len(temp) Then itempstr$ = TrimSpaces(itempstr$): GoTo itsdone
        If Len(TrimSpaces(itempstr$)) = 1 Then itempstr$ = "0" & TrimSpaces(itempstr$) 'If the number is a single digit then add a '0' to the front then hold it
        'Now check to see if a - number was created if so cause an error message
        If Left(TrimSpaces(itempstr$), 1) = "-" Then GoTo oops 'MsgBox "there is a bug in the ecryption method please contact whiteknight @ witenite87@excite.com with the following information." & vbCrLf & "Your Encryption Key, The Unencrypted String and the Encrypted String.", vbApplicationModal + vbCritical + vbDefaultButton1, "Encryption Error"
itsdone:
        temp2$ = temp2$ & itempstr 'hold the first decryption
        frmMain.txt_stp1.Text = frmMain.txt_stp1.Text & CInt(itempstr) + 100 & " "
    Next i%
       

    
    Encrypted = TrimSpaces(temp2$) 'set the encrypted data


    For i% = 1 To Len(Encrypted) Step 3
        'Format the encrypted string for the second decryption
        NewEncrypted = NewEncrypted & Mid(Encrypted, CLng(i%), 3) & " "
    Next i%
    lastTemp$ = TrimSpaces(Mid(NewEncrypted, Len(NewEncrypted$) - 3)) ' Hold the last set of numbers to check it its the correct format
    If Len(lastTemp$) = 2 Then ' If it = 2 then its not the Correct format and we need to fix it
        lastTemp$ = Mid(NewEncrypted, Len(NewEncrypted$) - 1) 'Holds Last Number so a '0' Can be added between them
        Encrypted = Mid(NewEncrypted, 1, Len(NewEncrypted) - 2) & "0" & lastTemp$ 'set it to the new format
    Else
        Encrypted$ = NewEncrypted$ 'set the new format
   frmMain.txt_stp2.Text = NewEncrypted$

    End If
    'The Actual Decryption
    For i% = 1 To Len(Encrypted)
        Letter = Mid$(Encrypted, i%, 1) 'Hold Letter at index i
        EncNum = EncNum & Letter 'Hold the letters
        If Letter = " " Then 'If letter =" " then we have a letter to decrypt
            encbuffer = CLng(Mid(EncNum, 1, Len(EncNum) - 1)) 'Convert it to long and get the number minus the " "
            strDecrypted$ = strDecrypted & Chr(encbuffer - Math) 'Store the decrypted string
            EncNum = "" 'clear if it is a space so we can get the next set of numbers
        End If
    Next i%
   
    Decrypt = strDecrypted
    frmMain.txt_stp3.Text = strDecrypted
     Decrypt = unscramb(strDecrypted$)
         frmMain.txt_stp4.Text = Decrypt
    Exit Function
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
    MsgBox "You Have Entered The WRONG Encrypt Key or Have the Worng Encrypted Message. " & vbCrLf & "If you think you received this Message in error please contact WhiteKnight @ witenite87@excite.com" & vbCrLf & "Please Include in the e-mail the following:" & vbCrLf & "The Encryption Key, The String To Encrypt and the Encrypted String.", vbApplicationModal + vbCritical + vbCritical + vbMsgBoxSetForeground, "Decryption Error"
End Function

Private Function TrimSpaces(strString As String) As String
    Dim lngpos As Long
    Do While InStr(1&, strString$, " ")
        DoEvents
        Let lngpos& = InStr(1&, strString$, " ")
        Let strString$ = Left$(strString$, (lngpos& - 1&)) & Right$(strString$, Len(strString$) - (lngpos& + Len(" ") - 1&))
    Loop
    Let TrimSpaces$ = strString$
End Function

Public Function scramb(strString As String) As String
Dim i As Integer, even As String, odd As String
For i% = 1 To Len(strString$)
If i% Mod 2 = 0 Then
even$ = even$ & Mid(strString$, i%, 1)
Else
odd$ = odd$ & Mid(strString$, i%, 1)
End If
Next i
scramb$ = even$ & odd$
End Function

Public Function unscramb(strString As String) As String
Dim x As Integer, evenint As Integer, oddint As Integer
Dim even As String, odd As String, fin As String
x% = Len(strString$)
x% = Int(Len(strString$) / 2) 'adding this returns the actuall number like 1.5 instead of returning 2
'Form1.Caption = x
even$ = Mid(strString$, 1, x%)
odd$ = Mid(strString$, x% + 1)
For x = 1 To Len(strString$)
If x% Mod 2 = 0 Then
evenint% = evenint% + 1
fin$ = fin$ & Mid(even$, evenint%, 1)
Else
oddint% = oddint% + 1
fin$ = fin$ & Mid(odd$, oddint%, 1)
End If
Next x%
unscramb$ = fin$
End Function


