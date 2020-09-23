VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Words"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   6840
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstWords 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11668
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Word"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Meaning"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   3600
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' www.stevenlees.net

' frmList
Option Explicit

Private WS As Workspace
Private DB As Database
Private RS As Recordset

Private i As Integer
Private j As Integer
Private strSQL As String

Private liWords As ListItem

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ' Save changes made.
    
    RS.MoveFirst
    
    For i = 1 To RS.RecordCount
    
        RS.Edit
    
        If lstWords.ListItems(i).Checked = True Then
            RS.Fields("use") = True
        Else
            RS.Fields("use") = False
        End If
        
        RS.Update
        RS.MoveNext
        
    Next
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    ' Load all words from the database.
    Set DB = OpenDatabase(App.Path & "\medicalterminology.mdb")
    
    strSQL = "SELECT * FROM tblWords WHERE type = """ & frmMain.strListSelect & """"
    
    Select Case frmMain.strListSelect
        Case "all": Set RS = DB.OpenRecordset("tblWords")
        Case Else: Set RS = DB.OpenRecordset(strSQL)
    End Select

    i = 1

    Do

        Set liWords = lstWords.ListItems.Add(i, , RS.Fields(0))
        
        If RS.Fields("use").Value = True Then
            lstWords.ListItems.Item(i).Checked = True
        End If
        
        For j = 1 To 3
            liWords.SubItems(j) = RS.Fields(j)
        Next
        
        i = i + 1
        RS.MoveNext

    Loop Until RS.EOF
    
    ListviewAutoFit lstWords, Label1

End Sub

' I did not write this function. It was found on PSC.
Public Sub ListviewAutoFit(ByRef List As ListView, ByRef AutosizeLabel As Label)
  'Automatically resizes listview columns so that all text
  'is visible. If the column name is empty, the column will
  'be resized to 0.
  'To make this procedure work properly, AutosizeLabel must
  'have the following settings:
  '   - Font has to be the same as the Listview's
  '   - Visible = False
  '   - AutoSize = True
  Dim i As Long
  Dim j As Long
  Dim State As Boolean
  
  With List
    State = .Visible
    .Visible = False
    
    For i = 1 To .ColumnHeaders.Count
      If .ColumnHeaders(i).Text <> "" Then
        AutosizeLabel.Caption = .ColumnHeaders(i).Text
        .ColumnHeaders(i).Width = AutosizeLabel.Width + 280
        For j = 1 To .ListItems.Count
          If i = 1 Then
            AutosizeLabel.Caption = .ListItems(j) _
                                  & IIf(.Icons Is Nothing, "", "XX")
          Else
            AutosizeLabel.Caption = .ListItems(j).SubItems(i - 1)
          End If
          
          If .ColumnHeaders(i).Width < AutosizeLabel.Width + 280 Then
            .ColumnHeaders(i).Width = AutosizeLabel.Width + 280
          End If
        Next
      Else
        .ColumnHeaders(i).Width = 0
      End If
    Next
    
    .Visible = State
  End With
  
  AutosizeLabel.Caption = "ID"
  lstWords.ColumnHeaders(1).Width = AutosizeLabel.Width + 500
  
  List.Visible = True
End Sub

