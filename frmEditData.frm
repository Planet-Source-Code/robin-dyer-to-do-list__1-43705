VERSION 5.00
Begin VB.Form frmEditData 
   Caption         =   "Edit Task"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   Icon            =   "frmEditData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDate 
      Height          =   300
      Left            =   1600
      TabIndex        =   10
      Top             =   2160
      Width           =   4000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   1575
      Left            =   1600
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   4000
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   2505
   End
   Begin VB.ComboBox cboRequest 
      Height          =   315
      Left            =   1600
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   " "
      Top             =   1200
      Width           =   2500
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   2505
   End
   Begin VB.TextBox txtDesc 
      Height          =   300
      Left            =   1600
      TabIndex        =   6
      Top             =   240
      Width           =   4000
   End
   Begin VB.Label Label6 
      Caption         =   "Comments:"
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Due Date:"
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Priority:"
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Requested By:"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim ii As Integer
    
    For ii = 2 To frmMain.cboStatus.ListCount 'Start at 2, to exclude "(All)" record
        cboStatus.AddItem frmMain.cboStatus.List(ii - 1)
    Next
    For ii = 2 To frmMain.cboPriority.ListCount
        cboPriority.AddItem frmMain.cboPriority.List(ii - 1)
    Next
    For ii = 2 To frmMain.cboRequest.ListCount
        cboRequest.AddItem frmMain.cboRequest.List(ii - 1)
    Next
    
    If iWhichRec > 0 Then
        'We're in "EDIT" mode, need to pull existing values from grid
        cboStatus = frmMain.MSFlexGrid1.TextMatrix(iWhichRec, 1)
        cboPriority = frmMain.MSFlexGrid1.TextMatrix(iWhichRec, 2)
        txtDesc = frmMain.MSFlexGrid1.TextMatrix(iWhichRec, 3)
        txtDate = frmMain.MSFlexGrid1.TextMatrix(iWhichRec, 4)
        cboRequest = frmMain.MSFlexGrid1.TextMatrix(iWhichRec, 5)
        txtComments = frmMain.MSFlexGrid1.TextMatrix(iWhichRec, 6)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim jj As Integer
    
    If iWhichRec > 0 Then
        'We're editing an existing task
        jj = iWhichRec
    Else
        'We're adding a new task
        lNumbTasks = lNumbTasks + 1
        frmMain.MSFlexGrid1.AddItem "TASK" & lNumbTasks
        jj = frmMain.MSFlexGrid1.Rows - 1
    End If
    frmMain.MSFlexGrid1.TextMatrix(jj, 1) = cboStatus.Text
    frmMain.MSFlexGrid1.TextMatrix(jj, 2) = cboPriority.Text
    frmMain.MSFlexGrid1.TextMatrix(jj, 3) = txtDesc.Text
    frmMain.MSFlexGrid1.TextMatrix(jj, 4) = txtDate.Text
    frmMain.MSFlexGrid1.TextMatrix(jj, 5) = cboRequest.Text
    frmMain.MSFlexGrid1.TextMatrix(jj, 6) = txtComments.Text
    
    frmMain.MSFlexGrid1.Row = jj
    HighPriorityColor frmMain.MSFlexGrid1
    If lNumbTasks > 0 Then
        AutosizeGridColumns frmMain.MSFlexGrid1, 15, 2000, Me
    End If

    frmMain.MSFlexGrid1.FixedRows = 1 'If last record has been deleted, this was set to zero. Reset.
    SaveData frmMain, frmMain.MSFlexGrid1
    Unload Me
End Sub
