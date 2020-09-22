VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CryptText"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChPass 
      Caption         =   "Change &Password"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "change password for Private Key"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenKey 
      Caption         =   "&Generate Key"
      Height          =   495
      Left            =   4500
      TabIndex        =   8
      ToolTipText     =   "Create Private Key"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdDwK 
      Caption         =   "De&crypt with Key"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      ToolTipText     =   "decrypt using private key"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdEwK 
      Caption         =   "E&ncrypt with Key"
      Height          =   495
      Left            =   1740
      TabIndex        =   6
      ToolTipText     =   "encrypt using private key"
      Top             =   6000
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   4500
      TabIndex        =   5
      ToolTipText     =   "open encrypted file"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sa&ve"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "save to a file"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Clo&se"
      Height          =   1095
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Height          =   495
      Left            =   1740
      TabIndex        =   2
      ToolTipText     =   "decrypt text"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "encrypt text"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '   CryptTextAES by Dipankar Basu
Private Sub cmdChPass_Click()
Dim myPassKey As String, currPass As String
    currPass = InputBox("Password for your key", "Input old password")
    If StrPtr(currPass) = 0 Then Exit Sub
    myPassKey = GetMyPassKey(currPass)
    If Len(myPassKey) = 25 Then  ' this is the length of PassPhrase, see GenKey(N)
        currPass = InputBox("Password for your key", "Input new password")
        If currPass = vbNullString Then Exit Sub
        myPassKey = Str2Hex(strEncrypt(myPassKey, currPass))
        SaveSetting "BasuDip Applications", "CryptTextAES", "PassPhraseKey", myPassKey
        myPassKey = Str2Hex(strEncrypt(currPass, StrReverse(currPass)))
        SaveSetting "BasuDip Applications", "CryptTextAES", "keyPassword", myPassKey
        MsgBox "Your password is successfully changed", , "Password Changed"
    Else
        MsgBox "The Password for your Key is Incorrect", , "Try again"
    End If
End Sub
Private Sub cmdDecrypt_Click()
Dim sTemp As String, sPassword As String
    sTemp = Trim$(Text1.Text)
    sPassword = InputBox("Input the password", "Decrypt with key", "Password Passphrase")
    If StrPtr(sPassword) = 0 Then Exit Sub
    sTemp = Hex2Str(sTemp)
    Text1.Text = strDecrypt(sTemp, sPassword)
End Sub
Private Sub cmdDwK_Click()
Dim sTemp As String, sPassword As String
    sTemp = Trim$(Text1.Text)
    sPassword = InputBox("PassPhrase Key Password", "Decrypt")
    If sPassword = vbNullString Then Exit Sub
    sPassword = GetMyPassKey(sPassword)
    If sPassword = vbNullString Then Exit Sub
    sTemp = Hex2Str(sTemp)
    Text1.Text = strDecrypt(sTemp, sPassword)
End Sub
Private Sub cmdEncrypt_Click()
Dim sTemp As String, sPassword As String
    sTemp = Text1.Text
    sPassword = InputBox("Input a password", "Encrypt with key", "Password Passphrase")
    If StrPtr(sPassword) = 0 Then Exit Sub
    sTemp = strEncrypt(sTemp, sPassword)
    Text1.Text = Str2Hex(sTemp)
End Sub
Private Sub cmdEwK_Click()
Dim sTemp As String, sPassword As String
    sTemp = Text1.Text
    sPassword = InputBox("PassPhrase Key Password", "Encrypt")
    If sPassword = vbNullString Then Exit Sub
    sPassword = GetMyPassKey(sPassword)
    If sPassword = vbNullString Then Exit Sub
    sTemp = strEncrypt(sTemp, sPassword)
    Text1.Text = Str2Hex(sTemp)
End Sub
Private Sub cmdExit_Click()
    Unload Me: End
End Sub
Private Sub cmdGenKey_Click()
Dim passPhrase As String, myKey As String, var As String, answer As Integer
    var = GetSetting("BasuDip Applications", "CryptTextAES", "PassPhraseKey")
    If var = vbNullString Then
        passPhrase = InputBox("Please provide a password to protect your key", "Create PassPhrase Private Key")
        If StrPtr(passPhrase) = 0 Or passPhrase = vbNullString Then Exit Sub
        myKey = Str2Hex(strEncrypt(GenKey(25), passPhrase))
        SaveSetting "BasuDip Applications", "CryptTextAES", "PassPhraseKey", myKey
        myKey = Str2Hex(strEncrypt(passPhrase, StrReverse(passPhrase)))
        SaveSetting "BasuDip Applications", "CryptTextAES", "keyPassword", myKey
        MsgBox "Password Key is Created, Please remember your password", vbInformation, "Create PassPhrase Key"
    Else
        answer = MsgBox("You are about to create a new key, any data that was encrypted with " & "your previous key becomes unrecoverable. You'll not be able to decrypt " & "any data encrypted with your old key." & vbNewLine & "Click on Change " & "Password to change the password for the existing key." & vbNewLine & "It is recommended to keep a backup of your old key, before creating a new key." & vbNewLine & "Do you wish to continue in creating a new PassPhrase key ?", vbYesNo + vbDefaultButton2 + vbCritical, "WARNING : Create a New Key")
        If answer = vbYes Then
            DeleteSetting "BasuDip Applications", "CryptTextAES", "PassPhraseKey"
            DeleteSetting "BasuDip Applications", "CryptTextAES", "keyPassword"
            Call cmdGenKey_Click
        End If
    End If
End Sub
Private Sub cmdOpen_Click()
Dim sFilename As String
    On Error GoTo eh:
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Open Encrypted File"
        .Filter = "Text Files(*.txt)|*.txt|Rich Text Files(*.rtf)|*.rtf|All Files(*.*)|*.*|"
        .InitDir = App.Path
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        sFilename = .FileName
    End With
    Open sFilename For Input As #1
    Text1.Text = StrConv(InputB$(LOF(1), 1), vbUnicode)
    Close #1
    If Right$(Text1.Text, 2) = vbNewLine Then Text1.Text = Left$(Text1.Text, Len(Text1.Text) - 2)
Exit Sub
eh:
    If Err.Number = 2755 Then Exit Sub ' CancelButton on OpenFile
End Sub
Private Sub cmdSave_Click()
Dim sFilePath As String, Fn As Integer
    sFilePath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
fnExists:
    Fn = Fn + 1
    If FileExists(sFilePath & "crypt" & Fn & ".txt") Then GoTo fnExists
    Open sFilePath & "crypt" & Fn & ".txt" For Output As #1
    Print #1, Text1.Text: Close #1
    MsgBox "File is saved to ..." & vbNewLine & sFilePath & "Crypt" & Fn & ".txt", vbInformation, "Text Encryption"
End Sub
Private Function GetMyPassKey(ByVal myPassword As String) As String
Dim regData As String, myKey As String
    If myPassword = vbNullString Then Exit Function
    myKey = Str2Hex(strEncrypt(myPassword, StrReverse(myPassword)))
    regData = GetSetting("BasuDip Applications", "CryptTextAES", "keyPassword")
    If myKey = regData Then
        regData = GetSetting("BasuDip Applications", "CryptTextAES", "PassPhraseKey")
        GetMyPassKey = strDecrypt(Hex2Str(regData), myPassword)
    Else
        GetMyPassKey = vbNullString
    End If
End Function
'Copyright (c)2003 by Dipankar Basu
' http://www.geocities.com/basudip_in/
