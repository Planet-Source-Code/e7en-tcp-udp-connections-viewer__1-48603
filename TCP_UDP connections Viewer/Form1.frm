VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "TCP/UDP Connection Viewer"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Cache Options"
      Height          =   700
      Left            =   1080
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5880
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ckResolve 
      Caption         =   "Resolve IP's"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CheckBox ckCache 
      Caption         =   "DNS Caching"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5950
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Protocol"
      Height          =   700
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   975
      Begin VB.ComboBox cPro 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":000A
         TabIndex        =   2
         Text            =   "TCP"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5950
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvNetstat 
      Height          =   3855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6800
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Protocol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'*************************************************************
'***                                                       ***
'***    Program Name: TCP/UDP Connection Viewer            ***
'***    Programmer:   Jake Paternoster (§e7eN)             ***
'***    Contact:      Hate_114@hotmail.com                 ***
'***    Date:         11:35 PM 18/09/2003                  ***
'***                                                       ***
'***    Description:  This program will show current       ***
'***                  connections for both TCP and UDP.    ***
'***                  Also Incorperated is a DNS class     ***
'***                  that will allow you to use a cache   ***
'***                  database to save time when resovling ***
'***                  muliple IP's. Its allows you to      ***
'***                  Import and Export lists aswell.      ***                                    ***
'***                                                       ***
'***                  Please Comment and vote              ***
'*************************************************************
'*************************************************************


Dim cLookup As clsLookUp

Private Sub cmdClose_Click()
    Unload Me                           'Unload the form
End Sub

Private Sub cmdExport_Click()
With CD
    .Filter = "Text Documents|*.txt"    'Set Dialogue box filter to only show text Documents
    .ShowSave                           'Show the save dialogue
    If .FileName = "" Then Exit Sub     'If no file selected then exit
    cLookup.ExportCache .FileName       'Save current cache to file
End With
End Sub

Private Sub cmdImport_Click()
With CD
    .Filter = "Text Documents|*.txt"    'Set Dialogue box filter to only show text Documents
    .ShowOpen                           'Show the Open dialogue
    If .FileName = "" Then Exit Sub     'If no file selected then exit
    cLookup.ImportCache .FileName       'Load file into cache
End With

End Sub

Private Sub cmdRefresh_Click()
    lvNetstat.ListItems.Clear   'Clear Previous Connections List
    GetConnections              'Get Current Connections
End Sub

Private Sub Form_Load()
Dim iTmp As Integer

Set cLookup = New clsLookUp     'create a new instance of the lookup class

'Resizing the columns in the list view box
For x = 1 To lvNetstat.ColumnHeaders.Count
    lvNetstat.ColumnHeaders(x).Width = Me.TextWidth(lvNetstat.ColumnHeaders(x).Text) * 1.3
    iTmp = iTmp + lvNetstat.ColumnHeaders(x).Width
Next

'Still resizeing
lvNetstat.ColumnHeaders(lvNetstat.ColumnHeaders.Count).Width = lvNetstat.Width - iTmp + 150

'Get Current connections
GetConnections
End Sub


Sub GetConnections()
Dim tUdpTable As MIB_UDPTABLE
Dim tTcpTable As MIB_TCPTABLE
Dim ldwSize As Long
Dim bOrder As Long

cLookup.DNScache = ckCache.value 'If DNS caching selected then enable

If cPro.Text = "UDP" Then        'If UDP connections selected then
'==============================================================================
'                               GET UDP CONNECTIONS
'==============================================================================

    Call GetUdpTable(tUdpTable, ldwSize, bOrder) 'Call it once to get ldwSize
    Call GetUdpTable(tUdpTable, ldwSize, bOrder)

    'cycle for every connection in the table
    For x = 0 To tUdpTable.dwNumEntries - 1
        'Add it to the info into the listview box
        lvNetstat.ListItems.Add , , "UDP"
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(1) = IPconvert(tUdpTable.table(x).dwLocalAddr)
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(2) = PortConvert(tUdpTable.table(x).dwLocalPort)
    Next
    
    If ckResolve.value = vbChecked Then 'If Resolving selected then
        'resolve all the IP's in the listview box
        For x = 1 To lvNetstat.ListItems.Count
            If lvNetstat.ListItems(x).SubItems(1) <> "0.0.0.0" Then lvNetstat.ListItems(x).SubItems(1) = cLookup.DNSlookup(lvNetstat.ListItems(x).SubItems(1))
        Next
    End If
    
Else                            'If TCP connections selected then

'==============================================================================
'                               GET TCP CONNECTIONS
'==============================================================================
'Very similar to above, so just read them comments

    Call GetTcpTable(tTcpTable, ldwSize, bOrder)
    Call GetTcpTable(tTcpTable, ldwSize, bOrder)

    For x = 0 To tTcpTable.dwNumEntries - 1
        lvNetstat.ListItems.Add , , "TCP"
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(1) = IPconvert(tTcpTable.table(x).dwLocalAddr)
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(2) = PortConvert(tTcpTable.table(x).dwLocalPort)
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(3) = IPconvert(tTcpTable.table(x).dwRemoteAddr)
        
        If tTcpTable.table(x).dwState = 2 Then
            lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(4) = ""
        Else
            lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(4) = PortConvert(tTcpTable.table(x).dwRemotePort)
        End If
        
        lvNetstat.ListItems(lvNetstat.ListItems.Count).SubItems(5) = StateConvert(tTcpTable.table(x).dwState)
    Next
    
    If ckResolve.value = vbChecked Then
        For x = 1 To lvNetstat.ListItems.Count
            If lvNetstat.ListItems(x).SubItems(1) <> "0.0.0.0" Then lvNetstat.ListItems(x).SubItems(1) = cLookup.DNSlookup(lvNetstat.ListItems(x).SubItems(1))
            If lvNetstat.ListItems(x).SubItems(3) <> "0.0.0.0" Then lvNetstat.ListItems(x).SubItems(3) = cLookup.DNSlookup(lvNetstat.ListItems(x).SubItems(3))
        Next
    End If
End If



End Sub

Private Sub lvNetstat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Sort the listview box depending on which column was clicked
lvNetstat.Sorted = True
lvNetstat.SortKey = ColumnHeader.Index - 1
End Sub

