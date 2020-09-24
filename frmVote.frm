VERSION 5.00
Begin VB.Form frmVote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thanks For Downloading This Code"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   4575
      Begin VB.CommandButton Ok 
         Caption         =   "&Go to Article Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1950
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Exit Application"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1950
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Label Label2 
         Caption         =   $"frmVote.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4365
      End
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const SW_SHOWNORMAL = 1
Const CodeID = 26900

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub Ok_Click()
    GotoURL ("http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=50695&lngWId=1")
End Sub



Sub GotoURL(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, Dum, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub


