VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "VBWeb Autocomplete example"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstParam 
      Height          =   645
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParameters 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   405
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.ListBox lstKeywords 
      Height          =   1230
      Left            =   4800
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtEdit 
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// VB Web Code Example
'// (c) 1999-2000
'// www.vbweb.co.uk/
'// You can do whatever you like with this code
'// IF you improve it/solve bugs
'// please email us with your modified code

'// THIS CODE MAY NOT BE USED IN A COMMERCIAL APPLICATION WITHOUT EXPLICIT PERMISSION
'// PLEASE DO NOT REDISTRIBUTE THIS CODE TO OTHER VB SITES
'// You may use the code in your own freeware programs without notifying VB Web
'// however, notification would be appreciated

Private Const LB_FINDSTRING = &H18F
Private Const LB_ERR = (-1)

Private Const EM_POSFROMCHAR = &HD6&
Private Const EM_LINEFROMCHAR = &HC9

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private strTotal As String
Private strPartial As String

Private lCurrentKeyWordIndex As Long
Private bEditFromCode As Boolean
Private bNoChange As Boolean

Public Function GetLastWord(Field As TextBox, lStart As Long)
    Dim lLine As Long
    Dim bIgnoringSpaces As Boolean
    If Len(Field.Text) < 1 Then Exit Function
    '// to avoid errors
    If lStart > 1 Then
        If Mid$(Field.Text, lStart - 1, 1) = "(" Or Mid$(Field.Text, lStart - 1, 1) = " " Then
            '// If we have an open bracket, go back to the
            '// last letter of the keyword
            lStart = lStart - 2
            If lStart > 1 Then
                Do While Mid$(Field.Text, lStart, 1) = " " And lStart > 1
                    lStart = lStart - 1
                Loop 'While lStart <> 1
                bIgnoringSpaces = True
            Else
                lStart = 1
            End If
            
        End If
    End If
    '// Find the last space from the current cursor point
    lLastSpace = InStrRev(Left$(Field.Text, lStart), " ")
    '// Find the last line from the current cursor point
    lLastLine = InStrRev(Left$(Field.Text, lStart), vbCrLf)
    
    If lLastSpace < lLastLine Then
        '// The last space is further away: use the last line
        lStartOfWord = lLastLine + 2
    Else
        '// the last line is further away: use the last space
        If lLastSpace = 1 Then lLastSpace = 0
        lStartOfWord = lLastSpace + 1
        
    End If
    '// if the start of word = 0, change it to 1,
    '// otherwise, we get an error!
    If lStartOfWord = 0 Then lStartOfWord = 1
    '// if we had to ignore the spaces, go forward
    '// one so that we do not chop of the last
    '// letter of the keyword
    If bIgnoringSpaces Then lStart = lStart + 1
    On Error Resume Next
    GetLastWord = Trim$(Mid$(Field.Text, lStartOfWord, lStart - lStartOfWord))
    If Err Then Debug.Print "ERROR!!"
End Function
Private Sub Form_Load()
    Dim vKeywords As Variant
    Dim vParameters As Variant
    Dim sKeywords As String
    Dim sParam As String
    '// load the keywords, seperated by *
    sKeywords = "MsgBox*InputBox*Mid*Left*"
    sParam = "Function MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context]) As VbMsgBoxResult*InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context]) As String*Function Mid(String, Start As Long, [Length])*Left(String, Length As Long)*"
    '// spit the strings into an array
    '// we are using a custom split function
    '// so that this code will work in VB5
    vKeywords = sSplit(sKeywords, "*")
    vParameters = sSplit(sParam, "*")
    For i = 0 To UBound(vKeywords)
        If vKeywords(i) <> "" Then
            '// load each of the keywords and parameters into
            '// the two list boxes
            lstKeywords.AddItem vKeywords(i)
            lstParam.AddItem vParameters(i)
        End If
    Next
End Sub

Private Sub Form_Resize()
    '// resize the text box
    txtEdit.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub lstKeywords_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTempPartial As String
    If KeyCode = vbKeySpace Then
        KeyCode = 0
        '// change the current word index
        lCurrentKeyWordIndex = lstKeywords.ListIndex
        '// change the actual word
        strTotal = lstKeywords.List(lstKeywords.ListIndex)
        '//
        strTempPartial = GetLastWord(txtEdit, txtEdit.SelStart + 1)
        '// we don't want the code in txtEdit_Change from firing
        bEditFromCode = True
        bNoChange = True
        '// remove the text that has already been entered
        txtEdit.SelStart = txtEdit.SelStart - Len(strTempPartial)
        txtEdit.SelLength = Len(strTempPartial)
        txtEdit.SelText = Left$(strTotal, 1)
        '// update the strPartial variable
        strPartial = Left$(strTotal, 1)
        '//
        txtEdit.SetFocus
        '// send the space bar message to the textbox
        SendKeys " ", True
        bEditFromCode = False
        bNoChange = False
    End If
End Sub

Private Sub txtEdit_Change()
    Dim i As Long, j As Long
    Dim xPixels As Long
    Dim yPixels As Long
    '// if this event has been triggered by our code, get out
    If bEditFromCode Then
        If Not bNoChange Then bEditFromCode = False
        Exit Sub
    End If
    '// hide the keyword listbox
    lstKeywords.Visible = False
    '// only allow valid keys
    With txtEdit
        '// get the last keyword from our text
        strPartial = GetLastWord(txtEdit, txtEdit.SelStart + 1)
        '// try to find a match on the keywords listbox
        i = SendMessage(lstKeywords.hwnd, LB_FINDSTRING, -1, ByVal strPartial)
        If i <> LB_ERR Then
            '// we have found a match
            '// select the match
            lstKeywords.ListIndex = i
            '// save it
            lCurrentKeyWordIndex = i
            '// get the full string
            strTotal = lstKeywords.List(i)
            '// if the keyword is not complete, show the list box
            If LCase$(strPartial) <> LCase$(strTotal) Then
                '// get the x,y co-ordinates
                GetPosFromChar .SelStart - 1, xPixels, yPixels
                '// move the keyword list to the
                '// correct position
                lstKeywords.Left = xPixels + txtEdit.Left
                lstKeywords.Top = yPixels + txtEdit.Top + TextHeight("Test")
                
                If lstKeywords.Top + lstKeywords.Height > ScaleHeight Then
                    lstKeywords.Top = yPixels - lstKeywords.Height
                End If
                lstKeywords.Visible = True
                txtParameters.Visible = False
            End If
        Else
            '// clear the strTotal variable
            If strPartial <> "(" Then strTotal = ""
        End If
    End With
End Sub

Private Sub txtEdit_Click()
    '// hide the parameters list if the line has changed (there
    '// is no selchange event with a text box)
    Call txtEdit_KeyUp(16, -1)
End Sub

Private Sub txtEdit_GotFocus()
    '// hide the parameters list if the line has changed (there
    '// is no selchange event with a text box)
    Call txtEdit_KeyUp(16, -1)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    '// ignore shift key only, for debugging
    If KeyCode = 16 Then Exit Sub
    Select Case Shift
    Case 0
        Select Case KeyCode
        Case vbKeySpace, vbKeyReturn, vbKeyBack, vbKeyDelete
            ShowParam (KeyCode)
        Case 187, 190 '// =.
            txtParameters.Visible = False
        Case 8
            txtParameters.Visible = False
        Case vbKeyUp
            If lstKeywords.Visible = True Then
                lstKeywords.SetFocus
                KeyCode = 0
                SendKeys "{Up}"
            End If
        Case vbKeyDown
            If lstKeywords.Visible = True Then
                lstKeywords.SetFocus
                KeyCode = 0
                SendKeys "{Down}"
            End If
        End Select
    Case vbShiftMask
        Select Case Chr(KeyCode)
        Case "9" '// (
            ShowParam (KeyCode)
        Case "0" '// )
            txtParameters.Visible = False
        End Select
    End Select
End Sub
Private Sub ShowParam(ByRef KeyCode As Integer)
    Dim xPixels As Long
    Dim yPixels As Long
    j = Len(strTotal) - Len(strPartial)
    Select Case KeyCode
    Case 41, 13, vbKeyBack, vbKeyDelete ' ), {Return}, {Back}
        If KeyCode = vbKeyBack Then
            If Mid$(txtEdit.Text, txtEdit.SelStart, 1) <> "(" Then
                On Error GoTo ErrSkip
                '// Gulp!
                
                If LCase$(Mid$(txtEdit.Text, txtEdit.SelStart - Len(strPartial) + 1, Len(strPartial))) <> _
                        LCase$(strPartial) And LCase$(Mid$(txtEdit.Text, txtEdit.SelStart - Len(strPartial), Len(strPartial))) _
                            <> LCase$(strPartial) Or strTotal = "" Then
                        
                    '// we not are eating into the keyword!
                    GoTo Continue
                End If
ErrSkip:
            End If
        End If
        txtParameters.Visible = False
        Exit Sub
    End Select
Continue:
    If strTotal <> "" Then
        If j <> 0 Then
            bEditFromCode = True
            '// complete the rest of the word
            txtEdit.SelText = Right$(strTotal, j)
        End If
        '// show the parameters label
        GetPosFromChar txtEdit.SelStart - 1, xPixels, yPixels
        txtParameters.Text = lstParam.List(lCurrentKeyWordIndex)
        txtParameters.Left = xPixels + txtEdit.Left
        txtParameters.Top = yPixels + txtEdit.Top + TextHeight("Test")
        txtParameters.Width = TextWidth(txtParameters.Text) + 10
        If txtParameters.Top + txtParameters.Height > ScaleHeight Then
            txtParameters.Top = yPixels - txtParameters.Height
        End If
                
        '// if the label is wider than the form
        '// wrap the text
        If txtParameters.Width > txtParameters.Left + ScaleWidth Then
            txtParameters.Height = 35
            txtParameters.Width = 313
        Else
            txtParameters.Height = 19
        End If
        lstKeywords.Visible = False
        txtParameters.Visible = True
        KeyCode = 0
    End If
End Sub
 Public Sub GetPosFromChar(ByVal lIndex As Long, ByRef xPixels As Long, ByRef yPixels As Long)
Dim lxy As Long
   lxy = SendMessageLong(txtEdit.hwnd, EM_POSFROMCHAR, lIndex, 0)
   xPixels = (lxy And &HFFFF&)
   yPixels = (lxy \ &H10000) And &HFFFF&
End Sub

Private Function sSplit(Expression As String, Optional Delimiter As String) As Variant
    Dim i As Long
    Dim lNextPos As Long
    Dim stext As String
    Dim lCount As Long
    Dim varTemp() As String
    For i = 1 To Len(Expression)
        lNextPos = InStr(i + 1, Expression, Delimiter)
        If lNextPos = 0 Then
            lNextPos = Len(Expression)
        End If
        stext = Mid$(Expression, i, lNextPos - i)
        ReDim Preserve varTemp(lCount)
        varTemp(lCount) = stext
        lCount = lCount + 1
        i = lNextPos
    Next
    sSplit = varTemp
End Function

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    Static lLastLine As Long
    Dim lLine As Long
    '// hide the parameters key if the line has changed
    lLine = SendMessage(txtEdit.hwnd, EM_LINEFROMCHAR, txtEdit.SelStart, 0&)
    If lLine <> lLastLine Then
        txtParameters.Visible = False
    End If
    lLastLine = lLine
    
    If KeyCode = 16 Then Exit Sub
    bValidItem = True
    
    Select Case Shift
    Case 0
    Case vbShiftMask
        Select Case Chr(KeyCode)
        Case "9" '// (
            '// show the parameters
            ShowParam (KeyCode)
        Case "0" '// )
            '// hide the parameters. A ) has been pressed
            txtParameters.Visible = False
        End Select
    End Select
End Sub

Private Sub txtParameters_Click()
    txtParameters.Visible = False
End Sub

Private Sub txtParameters_GotFocus()
    txtEdit.SetFocus
End Sub
