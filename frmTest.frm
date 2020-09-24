VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DB Tutorial"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "User Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   130
      Width           =   5415
      Begin VB.TextBox txtPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UserName :"
         Height          =   210
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PassWord:"
         Height          =   210
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5415
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "|< First"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "< Previous"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next >"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last >|"
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   2200
      Left            =   5640
      TabIndex        =   0
      Top             =   110
      Width           =   1455
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   610
         Width           =   1215
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Du&plicate"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Save"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "A&bandon"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   1720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Define Variables here
Dim cnAP As ADODB.Connection 'Connection to ADODB
Dim rsUsers As ADODB.Recordset 'Holds records
Dim mDuplicate As Boolean 'Variable to indicate if data has to be duplicated
Dim mNew As Boolean 'Variable to indicate if it's a new record
Dim mDirty As Boolean 'Variable to indicate if data is dirty or not
Dim mMove As Boolean 'To Find out if Data Changed or Traversed Recordset
Dim mBkMark As Variant 'Variable to Store BookMark
Dim bBookMarkable As Boolean 'Variable to indicate whether the recordset is bookmarkable or not
Public Sub DimControls(cString As String)
    'Procedure to Enable or Disable The Control Buttons (Add,Delete,Duplicate,Save,Abandon...
    Dim i As Integer
    Dim jcStr() As String
    jcStr = Split(cString, ",")
    For i = LBound(jcStr) To UBound(jcStr)
        cmdControl(i).Enabled = Val(jcStr(i))
    Next
End Sub
Private Sub cmdControl_Click(Index As Integer)
    'This module controls the click event of four command buttons
    'Add, Delete, Duplicate & Save
    If bBookMarkable And rsUsers.RecordCount > 0 Then
        mBkMark = rsUsers.Bookmark
    End If
    
    Dim strSQL As String
    Select Case Index
        Case 0 'Add New Record
            'New Record, Just Blank the Text boxes and set the focus on the first text box
            
            ClearControls
            mNew = True
            DimControls "0,0,0,1,1"
            DimNav "0,0,0,0"
        Case 1 'Delete Current Record
            'Delete! Just Confirm this so that you don't accidentally delete info
            If rsUsers.RecordCount > 0 Then
                If (MsgBox("Are You Sure You Want To Delete User : " & rsUsers!UserName, vbYesNo, "DB Turor") = vbYes) Then
                    With rsUsers
                        .Delete
                        .Requery
                        If EmptyDB(rsUsers) Then
                            Call ClearControls
                            mDirty = False
                        End If
                        LoadControls
                        Call cmdNavigate_Click(2)
                    End With
                End If
            End If
        Case 2 'Duplicate Record
            'Duplicate, Don't Clear the text boxes.
            'Wait for user to click SAVE.
            mDuplicate = True
            DimControls "0,0,0,1,1"
            DimNav "0,0,0,0"
            
        Case 3 'Save
            'Save.
            If mDuplicate Or mNew Then
                'Record is either new or duplicated.
                'So Insert it as a new record to the table
                strSQL = "INSERT INTO Users (UserName, Pwd) VALUES ('" _
                    & txtUserName.Text & "','" & txtPwd.Text & "');"
            Else
                'Record is edited
                'Update the existing record
                'Check for empty database
                If EmptyDB(rsUsers) Then
                    LoadControls
                    Exit Sub
                End If
                strSQL = "UPDATE Users SET UserName = '" & txtUserName.Text & "', Pwd = '" & txtPwd.Text & "'" _
                    & "WHERE Users.ID = " & rsUsers!id & ";"
            End If
            cnAP.Execute strSQL
            rsUsers.Requery
            mDuplicate = False
            mNew = False
            mDirty = False
            DimControls "1,1,1,0,0"
            DimNavX '
            Me.Caption = "DB Tester (" & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount & ")"

        Case 4 'Abandon
            If bBookMarkable And rsUsers.RecordCount > 0 Then
                rsUsers.Bookmark = mBkMark
            End If
            LoadControls
            mDuplicate = False
            mNew = False
            mDirty = False
        End Select
    mBkMark = -1
End Sub
Private Sub ClearControls()
    'This procedure clears the text boxes so that the user can type in new info
    'Called when the user click Add
    
    txtUserName.Text = vbNullString
    txtPwd.Text = vbNullString
    
End Sub
Private Sub cmdNavigate_Click(Index As Integer)
    'This is the navigational routine
    'When user clicks First, Previous, Next or Last buttons
    'the record pointer is moved accordingly
    'Navigational buttons are dimmed
    'and data is displayed on the form
    
    With rsUsers
        If EmptyDB(rsUsers) Then
            DimNavX
            Exit Sub
       End If
        Select Case Index
            Case 0
                .MoveFirst
            Case 1      'Previous
                .MovePrevious
                If .BOF Then
                    .MoveFirst
                End If
            Case 2      'Next
                .MoveNext
                If .EOF Then
                    .MoveLast
                End If
            Case 3      'Last
                .MoveLast
        End Select
        'Me.Caption = "DB Tester (" & .AbsolutePosition & " of " & .RecordCount & ")"
    End With
    mMove = True
    LoadControls
End Sub
Public Sub DimNav(cString As String)
    'Routine to disable / enable the navigational buttons
    'Same as dimcontrol
    'Both can be combined as a single routine,
    'taking cString and the control name as arguments.
    
    Dim jcStr() As String
    Dim i As Integer
    jcStr = Split(cString, ",")
    For i = LBound(jcStr) To UBound(jcStr)
        cmdNavigate(i).Enabled = Val(jcStr(i))
    Next
End Sub

Private Sub Form_Load()
    
    Set cnAP = New ADODB.Connection
    'Creates a new connection object
    Set rsUsers = New ADODB.Recordset
    'Creates a new Recordset object

    rsUsers.CursorLocation = adUseClient
    
    cnAP.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
           "Dbq=" & App.Path & "\mydb.mdb;" & _
           "Uid=admin;" & _
           "Pwd="
    'opens mydb.mdb which is located in the application path.
    cnAP.CursorLocation = adUseClient
    
    
    rsUsers.Open "SELECT * FROM Users;", cnAP, adOpenDynamic, adLockPessimistic
    'Opens the record set.
    
    LoadControls
    'Call the procedure to read data from recordset and display it in text boxes
    
    Me.Caption = "DB Tester (" & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount & ")"
    'Display the position and the total number of records in the Title Window
    
    bBookMarkable = IIf(rsUsers.Supports(adBookmark), True, False)
    'All recordsets does not support bookmarking, so make sure that
    'our recordset supports it, and if it does, then set bBookMarkable to True
    
    'We are using iif function here aswell (this has nothing to do with ADO though!)
    'iif is a short form of the if..else..endif block
    'the above statement can be written also as
    
    
'    If rsUsers.Supports(adBookmark) Then
'        bBookMarkable = True
'    Else
'        bBookMarkable = False
'    End If


End Sub
Private Sub LoadControls()
    With rsUsers
        If .RecordCount <= 1 Then
            DimNavX '"0,0,0,0"
        End If
        
        If EmptyDB(rsUsers) Then
            DimControls "1,0,0,0,0"
            Exit Sub
        End If
        If .BOF Then
            .MoveFirst
        End If
        If .EOF Then
            .MoveLast
        End If
        
        'This is the place where we read the information from recordset '
        'and store it to the variables.
               
        txtUserName.Text = !UserName & "" 'This assignment will err out if we
                                          ' try to assign a null value to the textbox
                                          'So an empty string is appended to it
        
        txtPwd.Text = !pwd & ""
        
    
    
    End With
    'The information is just displayed on the form and hence nothing is changed,
    'so We don't require Save & Abandon - disable it.
    DimNavX
    DimControls "1,1,1,0,0"
    
    
    'We check this variable when we unload the form to determine whether
    'any changes were made to the data.
    mDirty = False
        
        
    Me.Caption = "DB Tester (" & rsUsers.AbsolutePosition & " of " & rsUsers.RecordCount & ")"

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim jAnswer As Long
    If mDirty Then
        'This is where we check if the form is dirty (any data is changed).
        'If so, give the user a chance to save changes
        
        jAnswer = MsgBox("Do You Want To Save The Changes You Made To Users?", vbYesNoCancel + vbExclamation, "DB Tester")
        
        Select Case jAnswer
            Case vbYes
                Call cmdControl_Click(3)
                'User Clicked Yes
                'Call the procedure to save the information
                Cancel = False
            Case vbNo
                'User clicked No
                'Does not wish to save changes, So unload the form without saving
                Cancel = False
                
            Case vbCancel
                'User clicked cancel by mistake or changed mind to exit the app
                'so cancel unloading the form
                Cancel = True
                
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'This is where we do the clean up routine
    'Close all open connections
    'and set all objects to Nothing
    frmVote.Show
    
    rsUsers.Close
    cnAP.Close
    Set rsUsers = Nothing
    Set cnAP = Nothing
End Sub

Private Sub txtPwd_Change()
    'User changed data in the textbox, change the property Dirty to True
    If Not mMove Then
        DimControls "0,0,0,1,1"
        DimNav "0,0,0,0"
        mDirty = True
    End If
End Sub

Private Sub txtUserName_Change()
    'User changed data in the textbox, change the property Dirty to True
    If Not mMove Then
        DimControls "0,0,0,1,1"
        DimNav "0,0,0,0"
        mDirty = True
    End If
    
End Sub

Private Function EmptyDB(objrs As ADODB.Recordset) As Boolean
    If objrs.BOF And objrs.EOF Then
        EmptyDB = True
    Else
        EmptyDB = False
    End If
End Function
Private Sub DimNavX()
    'Enhancement of DimNav Procedure
    'Tested only with ClientSideCursors
    
    Dim jPos, jCount As Long
    With rsUsers
    
        jPos = .AbsolutePosition
        jCount = .RecordCount
    
    End With
    If jCount > 0 Then
        If jPos = 1 Then
            DimNav "0,0,1,1"
        Else
            If jPos = jCount Then
                DimNav "1,1,0,0"
            Else
                DimNav "1,1,1,1"
            End If
        End If
        If jCount = 1 Then
            DimNav "0,0,0,0"
        End If
    Else
        DimNav "0,0,0,0"
    End If
End Sub
