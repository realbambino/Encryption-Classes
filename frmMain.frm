VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Crypt"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbAlgorithm 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   900
      List            =   "frmMain.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   835
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   9900
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtOutputFileName 
      Height          =   285
      Left            =   900
      TabIndex        =   9
      Top             =   1215
      Width           =   8070
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   330
      Left            =   9090
      TabIndex        =   8
      Top             =   1215
      Width           =   1230
   End
   Begin VB.PictureBox picProgress 
      Height          =   330
      Left            =   855
      ScaleHeight     =   270
      ScaleWidth      =   9405
      TabIndex        =   6
      Top             =   1665
      Width           =   9465
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   900
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   485
      Width           =   3255
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   285
      Left            =   9090
      TabIndex        =   2
      Top             =   135
      Width           =   1230
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   135
      Width           =   8070
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Algorithm:"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   895
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Progress:"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1740
      Width           =   660
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   10350
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Out File:"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   530
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' High Encryption/Decryption - Made By Michael Ciurescu (CVMichael)

Option Explicit

Private Enum IfStringNotFound
    ReturnOriginalStr = 0
    ReturnEmptyStr = 1
End Enum

Private Enum eHashAlgorithm
    [MD5 - 128 Bit] = 1
    [SHA - 160 Bit] = 2
    [SHA256 - 256 Bit] = 3
End Enum

Private Type encFileHeader
    FLType1 As Byte ' = e
    FLType2 As Byte ' = n
    FLType3 As Byte ' = c
    Alg As Byte     ' algorithm = 1 or 2 or 3
    RndVal As Long  ' a random value (to enforce the password)
End Type

Private CancelCrypt As Boolean

Private Sub cmdBrowse_Click()
    On Error GoTo ErrCancel
    
    CDialog.Filter = "All Files (*.*)|*.*|Encrypted Files(*.enc)|*.enc"
    CDialog.ShowOpen
    
    txtFileName.Text = CDialog.FileName
    Exit Sub
ErrCancel:
    If Err.Number = 32755 Then
        Err.Clear
    Else
        MsgBox "Error: " & Err.Description, vbCritical, "Error# " & Err.Number
        Err.Clear
    End If
    Exit Sub
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Caption = "Start" Then
        cmdStart.Caption = "Cancel"
        
        CryptFile txtFileName.Text, txtOutputFileName.Text, txtPassword.Text, cmbAlgorithm.ListIndex + 1
    Else
        CancelCrypt = True
        cmdStart.Caption = "Start"
    End If
End Sub

Private Sub CryptFile(InFileName As String, OutFileName As String, Password As String, Optional Algorithm As eHashAlgorithm = [MD5 - 128 Bit])
    Dim InFileNum As Integer, OutFileNum As Integer, InBuff As String, OutBuff As String, PasswordCheck As String
    Dim FLHeader As encFileHeader, K As Long, Q As Long, HashLen As Long
    
    If Not (LCase(RightRight(InFileName, ".")) = "enc" Or LCase(RightRight(OutFileName, ".")) = "enc") Then Exit Sub
    
    picProgress.Scale (0, 0)-(1, 1)
    CancelCrypt = False
    
    InFileNum = FreeFile
    Open InFileName For Binary Access Read Lock Write As InFileNum
    
    If LCase(RightRight(InFileName, ".")) = "enc" Then ' decrypt
        Get InFileNum, , FLHeader
        Algorithm = FLHeader.Alg ' use same algorithm that was used to encrypt the file
        
        If FLHeader.FLType1 = Asc("e") And FLHeader.FLType2 = Asc("n") And FLHeader.FLType3 = Asc("c") And FLHeader.Alg >= 1 And FLHeader.Alg <= 3 Then
            HashLen = Choose(CInt(Algorithm), 16, 20, 32) ' if md5 then len = 16, sha = 20, sha256 = 32
            
            PasswordCheck = String(HashLen, 0)
            Get InFileNum, , PasswordCheck ' get hashed password from the encrypted file
            
            ' check if current hashed password is the same as the hashed password in the file
            If PasswordCheck <> HASH(Password & "?" & FLHeader.RndVal, False, Algorithm) Then
                MsgBox "Password incorrect, please try again"
                GoTo CloseAll
            End If
        Else
            MsgBox "Error, input file is not an encrypted file"
            GoTo CloseAll
        End If
        
        OutFileNum = FreeFile
        Open OutFileName For Binary Access Write Lock Write As OutFileNum
    Else ' encrypt
        OutFileNum = FreeFile
        Open OutFileName For Binary Access Write Lock Write As OutFileNum
        
        HashLen = Choose(CInt(Algorithm), 16, 20, 32) ' if md5 then len = 16, sha = 20, sha256 = 32
        
        Randomize
        FLHeader.FLType1 = Asc("e")
        FLHeader.FLType2 = Asc("n")
        FLHeader.FLType3 = Asc("c")
        FLHeader.Alg = CByte(Algorithm)
        FLHeader.RndVal = CLng(CLng(2 ^ 31 - 1) * Rnd) ' create a random value
        
        ' save file header
        Put OutFileNum, , FLHeader
        
        ' save hash of password with the random value (making it even more difficult to crack)
        ' also, when encrypting the same file 2 (or more) times, it will have different output
        ' every time (even if it's the same password)
        Put OutFileNum, , HASH(Password & "?" & FLHeader.RndVal, False, Algorithm)
    End If
    
    ' next portion of code will encrypt/decrypt the data from the file
    
    K = Not FLHeader.RndVal ' make the random number a negative number
    Do Until Loc(InFileNum) >= LOF(InFileNum) Or CancelCrypt
        If Loc(InFileNum) + HashLen > LOF(InFileNum) Then
            InBuff = String(LOF(InFileNum) - Loc(InFileNum), 0)
        Else
            InBuff = String(HashLen, 0)
        End If
        
        Get InFileNum, , InBuff
        
        ' get the hash string and XOR it with the input buffer, then save it into outputfile
        Put OutFileNum, , StrXOR(InBuff, Left(HASH(Password & "?" & K, False, Algorithm), Len(InBuff)))
        
        ' this ->  Password & "?" & K
        ' makes the hash output different all the time, therefore when the same character is repeated a lot
        ' it will always have a different output, having no patterns at all
        
        K = K + 1 ' increment K for the next buffer
        
        DoEvents
        picProgress.Line (0, 0)-(Loc(InFileNum) / LOF(InFileNum), 1), picProgress.ForeColor, BF
        picProgress.Line (Loc(InFileNum) / LOF(InFileNum), 1)-(1, 0), picProgress.BackColor, BF
    Loop
    
    txtPassword.Text = ""
    
CloseAll:
    Close InFileNum, OutFileNum
    
    cmdStart.Caption = "Start"
    picProgress.Line (0, 0)-(1, 1), picProgress.BackColor, BF
End Sub

Private Function StrXOR(Str1 As String, Str2 As String) As String
    Dim K As Integer
    
    If Len(Str1) <> Len(Str2) Then Exit Function
    
    StrXOR = String(Len(Str1), 0)
    For K = 1 To Len(Str1)
        Mid$(StrXOR, K, 1) = Chr(Asc(Mid$(Str1, K, 1)) Xor Asc(Mid$(Str2, K, 1)))
    Next K
End Function

Private Function HASH(Str As String, Optional ByVal ReturnHex As Boolean = True, Optional Algorithm As eHashAlgorithm = [MD5 - 128 Bit]) As String
    Dim Ret As String, K As Integer
    
    Select Case Algorithm
    Case [MD5 - 128 Bit]
        Dim cMD5 As New clsMD5
        
        Ret = cMD5.DigestStrToHexStr(Str)
        
        Set cMD5 = Nothing
    Case [SHA - 160 Bit]
        Dim cSHA As New clsSHA
        
        Ret = cSHA.SHA1(Str)
        
        Set cSHA = Nothing
    Case [SHA256 - 256 Bit]
        Dim cSHA256 As New clsSHA256
        
        Ret = cSHA256.SHA256(Str)
        
        Set cSHA256 = Nothing
    End Select
    
    If ReturnHex Then ' return hashed string as hex
        HASH = Ret
    Else ' return hashed string as binary
        HASH = String(Len(Ret) \ 2, 0)
        
        For K = 1 To Len(HASH)
            Mid$(HASH, K, 1) = Chr(Val("&H" & Mid$(Ret, K * 2, 2)))
        Next K
    End If
End Function

Private Sub Form_Load()
    cmbAlgorithm.ListIndex = 0
End Sub

Private Sub txtFileName_Change()
    cmdStart.Enabled = Dir(txtFileName.Text, vbArchive + vbHidden) <> "" And Trim(txtFileName.Text) <> ""
    
    If LCase(RightRight(txtFileName.Text, ".")) = "enc" Then
        txtOutputFileName.Text = RightLeft(txtFileName.Text, ".")
        cmbAlgorithm.Enabled = False
    Else
        txtOutputFileName.Text = txtFileName.Text & ".enc"
        cmbAlgorithm.Enabled = True
    End If
End Sub

' Search from end to beginning, and return the left side of the string
Private Function RightLeft(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long
    
    K = InStrRev(Str, RFind, , Compare)
    
    If K = 0 Then
        RightLeft = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightLeft = Left(Str, K - 1)
    End If
End Function

' Search from end to beginning and return the right side of the string
Private Function RightRight(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String
    Dim K As Long
    
    K = InStrRev(Str, RFind, , Compare)
    
    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightRight = Mid(Str, K + 1, Len(Str))
    End If
End Function
