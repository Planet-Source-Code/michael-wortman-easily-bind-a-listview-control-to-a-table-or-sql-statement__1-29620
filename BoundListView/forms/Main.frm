VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Populate"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim rst As ADODB.Recordset
    Dim db As ADODB.Connection
    
    ' initialize active data objects
    Set db = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    ' open db and then People table
    db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\test.mdb;Persist Security Info=False")
    Call rst.Open("People", db, adOpenDynamic)
    
    ' populate listview control
    Call PopulateList(ListView1, rst)
End Sub

Private Sub PopulateList(pList As ListView, _
                         pRst As ADODB.Recordset)
    On Error Resume Next
    Dim i As Long
    Dim iColCount As Long
    Dim sColName As String
    Dim sColValue As String
    Dim oCH As ColumnHeader
    Dim oLI As ListItem
    Dim oSI As ListSubItem
    Dim oFld As ADODB.Field


    With pList
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .ListItems.Clear
        .Sorted = False
        pRst.MoveFirst
        ' set up column headers


        For Each oFld In pRst.Fields
            sColName = cn(oFld.Name)
            Set oCH = .ColumnHeaders.Add()
            oCH.Text = sColName
            iColCount = iColCount + 1
        Next oFld


        Do Until pRst.EOF
            i = 0
            ' setup fiprst column as a listitem
            sColValue = cn(pRst.Fields(i).Value)
            Set oLI = .ListItems.Add()
            oLI.Text = sColValue
            ' add the remaining columns
            'as subitems


            For i = 1 To iColCount
                Set oSI = oLI.ListSubItems.Add()
                oSI.Text = cn(pRst(i))
            Next ' next column
            pRst.MoveNext
        Loop ' Next record
        ' refresh it all
        .Refresh
        ' make sure 1st row can be seen
        .ListItems(1).EnsureVisible
    End With
End Sub

' Function cn "Catch Null"
'   Returns a blank string if a null value is encountered
'
' this function makes sure that if the database returns
'   a null value for a field that we output an empty string
'   to the listview.  Otherwise, if you try to put a null
'   value to a listview (or any) control you will generate
'   an error
Private Function cn(pVal As String)
    If IsMissing(pVal) Then
        cn = ""
    ElseIf IsNull(pVal) Then
        cn = ""
    Else
        cn = Format(pVal)
    End If
End Function
