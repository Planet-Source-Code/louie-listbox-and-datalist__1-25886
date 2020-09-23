VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Caption         =   "Listbox and Datalist Example"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Re&move Duplicated"
      Height          =   975
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2280
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   4155
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7329
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   " R&emove All ->"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   " &Remove ->"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<- A&dd All"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<- &Add"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This example will show you how to move items _
'from datalist (which is connected to ADO control) to a listbox.
'This is not by far a perfect example, If you _
'like to send me any comments Email me at: babylon3000@hotmail.com.

Private Sub Command1_Click()
On Error Resume Next
If DataList1.VisibleCount = -1 Then
MsgBox "Nothing to Add?", vbInformation, "Error"
Else
List1.AddItem DataList1.Text
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command2_Click()
Dim X
On Error Resume Next
If DataList1.VisibleCount = -1 Then
MsgBox "Nothing to Add?", vbInformation, "Error"
Else
For X = 0 To DataList1.VisibleCount - 1
List1.AddItem DataList1.Text
Adodc1.Recordset.MoveNext
Next X
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List1.ListCount = -1 Then
MsgBox "Nothing to remove?", vbInformation, "Error"
Else
Dim P As Integer
P = List1.ListIndex
List1.RemoveItem (P)
End If
End Sub

Private Sub Command4_Click()
List1.Clear
End Sub

Private Sub Command5_Click()
Dim X As Integer
X = 0
Do While X < List1.ListCount
    List1.Text = List1.List(X)
    If List1.ListIndex <> X Then
        List1.RemoveItem X
    Else
        X = X + 1
    End If
Loop
End Sub

Private Sub Form_Load()
Dim ConnectionString As String
   Dim strDatabase As String
   strDatabase = App.Path & "\datalist.mdb"
ConnectionString = "Provider =" & "Microsoft.Jet.OLEDB.3.51;Data Source= " & strDatabase

   With Adodc1
      .RecordSource = _
      "SELECT CUS_ID, Email FROM Customers"
      .ConnectionString = ConnectionString
      .Caption = "Customers"
      .Refresh
      .Visible = False
   End With
   
   With DataList1
      Set .DataSource = Adodc1
      .DataField = "Email"
      .BoundColumn = "Email"
      Set .RowSource = Adodc1
      .ListField = "Email"
   End With
   
   Adodc1.Recordset.MoveFirst
End Sub
