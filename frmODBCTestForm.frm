VERSION 5.00
Begin VB.Form frmODBCTest 
   Caption         =   "ODBC Test Project"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetSVR 
      Caption         =   "Get Server"
      Height          =   345
      Left            =   2550
      TabIndex        =   15
      Top             =   3450
      Width           =   1260
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   1260
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2850
      Width           =   2700
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1260
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2400
      Width           =   2700
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1260
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1950
      Width           =   2700
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1260
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1500
      Width           =   2700
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1020
      Width           =   2700
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1260
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   540
      Width           =   2700
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   75
      Width           =   2700
   End
   Begin VB.CommandButton cmdSetDSN 
      Caption         =   "Set DSN"
      Height          =   345
      Left            =   900
      TabIndex        =   0
      Top             =   3450
      Width           =   1260
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Login ID:"
      Height          =   195
      Left            =   465
      TabIndex        =   13
      Top             =   2460
      Width           =   645
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Driver Path:"
      Height          =   195
      Left            =   270
      TabIndex        =   12
      Top             =   2010
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Driver Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   270
      TabIndex        =   10
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DSN:"
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DB Name:"
      Height          =   195
      Left            =   375
      TabIndex        =   8
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "frmODBCTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function SetDSN(ByVal DB_Name As String, _
                        ByVal DSN As String, _
                        ByVal Description As String, _
                        ByVal Driver_Name As String, _
                        ByVal Driver_Path As String, _
                        ByVal Last_User As String, _
                        ByVal Server_Name As String, _
                        ByRef Status As String _
                        ) As Boolean

   Dim ThisODBC As New clsODBC
   Dim msg As String
   
   
   'the name of the DB (i.e. NorthWind)
   ThisODBC.DatabaseName = DB_Name
   
   'the name you want to appear in Control Panel|ODBC Data Source Administrator|System DSN
   ThisODBC.DataSourceName = DSN
   
   'the description as it appears in Control Panel|ODBC Data Source Administrator|DSN Configuration
   ThisODBC.Description = Description
   
   'the driver to be used for the data source
   ThisODBC.DriverName = Driver_Name
   
   'path to the SQL Server driver
   ThisODBC.DriverPath = Driver_Path
   
   'SQL Server Sys Admin Login Name
   ThisODBC.LastUser = Last_User
   
   'computer name where the SQL Server DB lives (note: "\\" is not needed in front of name)
   ThisODBC.Server = Server_Name
   
   'given these attributes of an ODBC connection, SetDSN will look to see
   'if it exists exactly as given, and if it doesnt, will either Add or Update it.
   'Either way, if it returns true then the DSN will be set to the properties given....
   SetDSN = ThisODBC.SetDSN
   
   'bring the status back out of the object
   Status = ThisODBC.Status
   
   Set ThisODBC = Nothing
    
End Function

Private Function GetServer(ByVal ThisDSN As String, ByRef ThisServer As String, ByRef ThisStatus As String) As Boolean

   Dim ThisODBC As New clsODBC
   
   'set the one you want to find out about
   ThisODBC.DataSourceName = ThisDSN
   
   'given the datasourcename GetServer will return the server
   'that is assigned to it in ODBC.
   GetServer = ThisODBC.GetServer
   
   'set the var to return the server name
   ThisServer = ThisODBC.Server
   
   'set the status to return
   ThisStatus = ThisODBC.Status
   
   Set ThisODBC = Nothing
    
End Function

Private Sub cmdGetSVR_Click()

   Dim msg As String
   Dim ThisSVRName As String
   
   msg = ""
   
   If GetServer(Text2.Text, ThisSVRName, msg) _
   Then
      msg = msg & vbCrLf & vbCrLf & "Server Name: " & ThisSVRName & vbCrLf
      MsgBox msg, vbExclamation, "ODBC Test Project"
   Else
      MsgBox msg, vbCritical, "ODBC Test Project"
   End If


End Sub

Private Sub cmdSetDSN_Click()

   Dim msg As String
   
   msg = ""
   
   If SetDSN(Text1.Text, Text2.Text, Text3.Text, _
             Text4.Text, Text5.Text, Text6.Text, _
             Text7.Text, msg) _
   Then
      MsgBox msg, vbExclamation, "ODBC Test Project"
   Else
      MsgBox msg, vbCritical, "ODBC Test Project"
   End If


End Sub

Private Sub Form_Load()

   Text1.Text = "MyDB"
   
   Text2.Text = "My New DSN"
   
   Text3.Text = "ODBC Test Project"
   
   Text4.Text = "SQL Server"
   
   Text5.Text = "C:\Winnt\System32\sqlsrv32.dll"
   
   Text6.Text = "sa"
   
   Text7.Text = "SQLDEVSVR"


End Sub
