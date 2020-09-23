VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMain 
   Caption         =   "To Do List"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   10215
      Begin VB.CheckBox chkHideComplete 
         Caption         =   "Hide Completed Tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton cmdNewTask 
         Caption         =   "New Task"
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditTask 
         Caption         =   "Edit Task"
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelTask 
         Caption         =   "Delete Task"
         Height          =   495
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   8520
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Text            =   "(All)"
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Text            =   "(All)"
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboRequest 
         Height          =   315
         Left            =   8280
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "(All)"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter Tasks by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Requested By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboPriority_Click()
    FilterReport Me.MSFlexGrid1
End Sub

Private Sub cboRequest_Click()
    FilterReport Me.MSFlexGrid1
End Sub

Private Sub cboStatus_Click()
    FilterReport Me.MSFlexGrid1
End Sub

Private Sub chkHideComplete_Click()
    FilterReport Me.MSFlexGrid1
End Sub

Private Sub cmdDelTask_Click()
    Dim ii As Integer
    ii = MSFlexGrid1.Row
    If ii > 0 And ii <= lNumbTasks Then
        Select Case MsgBox("Are you sure you want to delete this Task?", vbYesNo + vbQuestion + vbDefaultButton1, App.Title)
            Case vbYes
                GridDeleteRow frmMain.MSFlexGrid1
                lNumbTasks = lNumbTasks - 1
                DeleteINI_Section myTaskFile, "TASK" & lNumbTasks + 1 'Get rid of last record, so remainder can be over-written
                RenumberTasks frmMain.MSFlexGrid1
                SaveData Me, MSFlexGrid1
            Case vbNo
                ' Exit, Do Nothing
        End Select
    End If
End Sub

Private Sub cmdEditTask_Click()
    iWhichRec = MSFlexGrid1.Row
    frmEditData.Show
End Sub

Private Sub cmdExit_Click()
    Dim Form As Form
    For Each Form In Forms
        Unload Form
    Next Form
End Sub

Private Sub cmdNewTask_Click()
    iWhichRec = 0
    frmEditData.Show
End Sub

Private Sub Form_Load()
    InitVars
    LoadData Me, MSFlexGrid1
    LoadRequestedBy Me, MSFlexGrid1
    MSFlexGrid1.SelectionMode = flexSelectionByRow
End Sub

Private Sub Form_Resize()
    If Me.Height > 800 And Me.Width > 800 Then
        MSFlexGrid1.Width = Me.Width - 400
        MSFlexGrid1.Height = Me.Height - 2670
        Frame1.Top = Me.Height - 2445
        Frame1.Width = Me.Width - 400
        If lNumbTasks > 0 Then
            AutosizeGridColumns MSFlexGrid1, 15, 2000, Me
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWindowPosition
End Sub

Private Sub MSFlexGrid1_Click()
    If MSFlexGrid1.MouseRow = 0 Then
        SortGrid Me.MSFlexGrid1, MSFlexGrid1.MouseCol
    End If
End Sub

Private Sub MSFlexGrid1_DblClick()
    iWhichRec = MSFlexGrid1.Row
    If iWhichRec > 0 Then
        frmEditData.Show
    End If
End Sub
