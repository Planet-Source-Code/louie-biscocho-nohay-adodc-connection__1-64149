VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simple Phone Book"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   120
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   3840
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================
' Codes by   :    Louie Biscocho Nohay
' Email Add:      lbnohay@yahoo.com
' Website:        www.noborsoft.cjb.net
' Tel. No.   :    +63.43.984.8338
' ======================================

Option Explicit

Dim i As Integer

Private Sub Command2_Click()
If Command2.Caption = "&Add" Then
   Call Clear_Text(Text2)
   Text2(0).SetFocus
Else
   Text2(0).SetFocus
End If
SetButton False
End Sub

Private Sub Command3_Click()
For i = 0 To Text2.Count - 1
   If Text2(i).Text = "" Then
      Inform "Fields are empty!" & vbCrLf & "Please fill in required fields", "Information"
      Exit Sub
   End If
Next i

With Adodc1.Recordset
If Command3.Caption = "&Save" Then
   .AddNew
   .Fields(0) = (Text2(0).Text)
   .Fields(1) = (Text2(1).Text)
   .Fields(2) = (Text2(2).Text)
   .Update
   Inform "New record has been added!", "Information"
   Call Clear_Text(Text2)
End If
If Command3.Caption = "&Update" Then
   .Fields(0) = (Text2(0).Text)
   .Fields(1) = (Text2(1).Text)
   .Fields(2) = (Text2(2).Text)
   .Update
   Inform "Record has been updated!", "Information"
   Call Clear_Text(Text2)
End If
End With
SetButton True
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
SetButton True
Call Clear_Text(Text2)
End Sub

Private Sub Command5_Click()
Adodc1.Refresh
Call Clear_Text(Text2)
Command2.Caption = "&Add"
Command3.Caption = "&Save"
SetButton True
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Cancel
Call Clear_Text(Text2)
Command2.Caption = "&Add"
Command3.Caption = "&Save"
SetButton True
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
With Adodc1.Recordset
   Text2(0).Text = .Fields(0)
   Text2(1).Text = .Fields(1)
   Text2(2).Text = .Fields(2)
End With
End Sub

Private Sub Form_Load()
'get db connection
Inform "This app basically shows connection using ADO Data Control in an organized way.", "Simple Phone Book"
Call get_connection(Adodc1, "Select * From Table1 Order By Name", App.Path & "\db1.mdb", False, "")

' initialize datagrid
With DataGrid1
   Set .DataSource = Adodc1
   .AllowUpdate = False
End With

Call Clear_Text(Text2)
End Sub

' clear textboxes
Private Sub Clear_Text(sTextBox As Object)
Dim ctr As Integer

For ctr = 0 To sTextBox.Count - 1
   sTextBox(ctr).Text = ""
Next ctr

End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Recordset.Close
Inform "Program Codes By: Louie Biscocho Nohay" & vbCrLf & vbCrLf & "Website:           www.noborsoft.cjb.net" & vbCrLf & _
        "Email Address:   lbnohay@yahoo.com" & vbCrLf & "Contact No.      +63.43.984.8338" & vbCrLf & vbCrLf & _
        "Please don't forget to vote for this app! Thanks!", "Simple Phone Book"
End Sub

Private Sub Text2_Change(Index As Integer)
On Error Resume Next
With Adodc1.Recordset
   If .RecordCount < 1 Then Exit Sub
   .MoveFirst
   .Find "Name Like '" & Text2(0).Text & "'"
   If .EOF Then
      Exit Sub
   Else
      Text2(1).Text = .Fields(1)
      Text2(2).Text = .Fields(2)
      Command3.Caption = "&Update"
      Command2.Caption = "&Edit"
    End If
End With
End Sub

Private Sub Text2_GotFocus(Index As Integer)
For i = 0 To Text2.Count - 1
   Call HyLyt(Text2(i))
Next i
End Sub

Sub SetButton(TF As Boolean)
Command2.Enabled = TF
Command3.Enabled = Not TF
End Sub
