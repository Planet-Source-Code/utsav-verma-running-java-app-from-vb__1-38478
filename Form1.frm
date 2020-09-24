VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Run java"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Set Bin Path"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "File"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Cdo 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compile"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "File Name"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Java bin directory"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If Trim(Text2.Text) <> "" Or Trim(Text1.Text) = "" Then

    Dim prg As String, dd As String, arg As String, tmp() As String
    
    prg = Text1.Text & "\javac.exe" & vbNullString
    
    tmp = Split(Text2.Text, "\")
    
    For i = 0 To UBound(tmp)
        
        If i = UBound(tmp) Then
            arg = tmp(i)
        Else
            dd = dd & tmp(i) & "\"
        End If
        
    Next
    
    SetCurrentDirectory dd
    
    MsgBox ShellExecute(Me.hwnd, "open", prg, arg, dd, SW_SHOWNA)
Else
    MsgBox "specify file name/path"
End If
End Sub

Private Sub Command2_Click()
If Trim(Text2.Text) <> "" Or Trim(Text1.Text) = "" Then
    Dim prg As String, dd As String, arg As String, tmp() As String
    Dim x As String
    Dim i As Integer
    
    x = Space(250)
    prg = Text1.Text & "\java.exe" & vbNullString
    
    tmp = Split(Text2.Text, "\")
    
    For i = 0 To UBound(tmp)
        
        If i = UBound(tmp) Then
            arg = tmp(i)
            'cut off '.java' portion
            arg = Mid$(arg, 1, Len(arg) - 5)
        Else
            dd = dd & tmp(i) & "\"
        End If
    Next
    
    SetCurrentDirectory dd
        
    ShellExecute Me.hwnd, "open", prg, arg, dd, SW_SHOW

Else
    MsgBox "specify file name/path"
End If
End Sub

Private Sub Command4_Click()
Text1.Text = GetFolder("Select Folder which has to be uploaded", Me.hwnd)
SaveSetting "run java", "path", "bin", Text1
End Sub

Private Sub Text1_GotFocus()

If Trim(GetSetting("run java", "path", "bin")) = "" Then
    Text1.Text = GetFolder("Select Folder which has to be uploaded", Me.hwnd)
    SaveSetting "run java", "path", "bin", Text1
Else
    Text1.Text = GetSetting("run java", "path", "bin")
End If
End Sub

Private Sub Text2_GotFocus()
Cdo.ShowOpen
Text2.Text = Cdo.FileName
End Sub
