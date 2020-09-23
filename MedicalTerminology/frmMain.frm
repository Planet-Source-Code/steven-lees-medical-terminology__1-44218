VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medical Terminology"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin HTTSLibCtl.TextToSpeech ttsMain 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":0442
      TabIndex        =   25
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Options"
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   1935
      Begin VB.CheckBox chkSpeakWord 
         Caption         =   "Speak Word"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkShowType 
         Caption         =   "Show Type"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkShowMeaning 
         Caption         =   "Show Meaning"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkRandomize 
         Caption         =   "Randomize"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkPopUp 
         Caption         =   "Pop-Up Meaning"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame frame 
      Height          =   4095
      Left            =   2160
      TabIndex        =   19
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optStatus 
         Caption         =   "Disable this word"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Enable this word"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtIndex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtMeaning 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtWord 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Index:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblMeaning 
         Caption         =   "Meaning:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblWord 
         Caption         =   "Word:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame frameWordTypes 
      Caption         =   "Word Types"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optWordType 
         Caption         =   "All Word Types"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optWordType 
         Caption         =   "Roots"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optWordType 
         Caption         =   "Suffixes"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optWordType 
         Caption         =   "Prefixes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   5775
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   16  'Merge Pen
      X1              =   120
      X2              =   5760
      Y1              =   4335
      Y2              =   4335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShowAllWords 
         Caption         =   "&Show All Words"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPrefixes 
         Caption         =   "Show &Prefixes"
      End
      Begin VB.Menu mnuShowSuffixes 
         Caption         =   "Show &Suffixes"
      End
      Begin VB.Menu mnuShowRoots 
         Caption         =   "Show &Roots"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' www.stevenlees.net

' frmMain
Option Explicit

Private WS As Workspace
Private DB As Database
Private RS As Recordset

Private i As Integer
Private j As Integer
Private strSQL As String

Public strListSelect As String
Private strType As String
Private intCurrentWord As Integer
Private strMessage As String
Private intPrevious As Integer
Private intLastWord As Integer
Private intLastWordFirst As Integer
Private intLastWordMax As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

Redo:

    cmdPrevious.Enabled = True
    
    intPrevious = txtIndex.Text
    
    txtType.Text = ""
    txtMeaning.Text = ""

    Select Case strType
        Case "all" ' Load all the words.
        
            Set RS = DB.OpenRecordset("tblWords")
            
            ' Get a word.
            If chkRandomize.Value = 1 Then
                strSQL = "SELECT * FROM tblWords WHERE id = " & RandomRange(1, RS.RecordCount)
            Else
                strSQL = "SELECT * FROM tblWords WHERE id = " & intLastWord + 1
            End If
            
        Case Else
        
            ' Get words of specific type.
            strSQL = "SELECT * FROM tblWords WHERE type = """ & strType & """"
            Set RS = DB.OpenRecordset(strSQL)
      
            j = RS.Fields("id")
            
            i = 0
            Do
                i = i + 1
                RS.MoveNext
            Loop Until RS.EOF

            If chkRandomize.Value = 1 Then
                strSQL = "SELECT * FROM tblWords WHERE type = """ & strType & """ AND id = " & RandomRange(j, i)
            Else
                If intLastWord = intLastWordMax Then intLastWord = intLastWordFirst - 1
                strSQL = "SELECT * FROM tblWords WHERE type = """ & strType & """ AND id = " & intLastWord + 1
            End If
            
    End Select
    
    Set RS = DB.OpenRecordset(strSQL)
        
    ' Skip words that aren't in use.
    If RS.Fields("use").Value = False Then
        intLastWord = intLastWord + 1
        GoTo Redo
    End If

    txtWord.Text = Trim(RS.Fields("word"))
    txtIndex.Text = Trim(RS.Fields("id"))
    
    If chkSpeakWord.Value = 1 Then ttsMain.Speak (txtWord.Text)
    
    intLastWord = txtIndex.Text
    
    If chkShowMeaning.Value = 1 Then txtMeaning.Text = Trim(RS.Fields("meaning"))
    If chkShowType.Value = 1 Then txtType.Text = Trim(RS.Fields("type"))
    
    If RS.Fields("use").Value = True Then optStatus(0).Value = True
    If RS.Fields("use").Value = False Then optStatus(1).Value = True

End Sub

Private Sub cmdPrevious_Click()
    ' Get the previous word.
    
    cmdPrevious.Enabled = False
    
    If chkRandomize.Value = 0 Then intLastWord = intLastWord - 1
    
    txtType.Text = ""
    txtMeaning.Text = ""
    
    strSQL = "SELECT * FROM tblWords WHERE id = " & intPrevious
    
    Set RS = DB.OpenRecordset(strSQL)
    
    txtWord.Text = Trim(RS.Fields("word"))
    txtIndex.Text = Trim(RS.Fields("id"))
    
    If chkSpeakWord.Value = 1 Then ttsMain.Speak (txtWord.Text)
    
    If chkShowMeaning.Value = 1 Then txtMeaning.Text = Trim(RS.Fields("meaning"))
    If chkShowType.Value = 1 Then txtType.Text = Trim(RS.Fields("type"))
    
    If RS.Fields("use").Value = True Then optStatus(0).Value = True
    If RS.Fields("use").Value = False Then optStatus(1).Value = True
    
    cmdNext.SetFocus

End Sub

Private Sub cmdShow_Click()
    ' Show meaning and type.
    
    txtType.Text = Trim(RS.Fields("type"))
    txtMeaning.Text = Trim(RS.Fields("meaning"))
    
    If chkSpeakWord.Value = 1 Then ttsMain.Speak (txtMeaning.Text)
    
    If chkPopUp.Value = 1 Then strMessage = MsgBox(txtMeaning.Text, vbOKOnly + vbInformation, "Medical Terminology")
    
End Sub

Private Sub Form_Load()
    
    ' Load database. Must be in app's home directory.
    Set DB = OpenDatabase(App.Path & "\medicalterminology.mdb")
    Set RS = DB.OpenRecordset("tblWords")

    ' All types - prefixes, suffixes and roots
    strType = "all"
    
    ' Get a random word.
    strSQL = "SELECT * FROM tblWords WHERE id = " & RandomRange(1, RS.RecordCount)
    Set RS = DB.OpenRecordset(strSQL)
    
    txtWord.Text = Trim(RS.Fields("word"))
    txtIndex.Text = Trim(RS.Fields("id"))
    
    ' Set the text-to-voice voice speed.
    ttsMain.Speed = 150
    
    ' Speak the word.
    If chkSpeakWord.Value = 1 Then ttsMain.Speak (txtWord.Text)
    
    intPrevious = txtIndex.Text
    intLastWord = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuShowAllWords_Click()
    strListSelect = "all"
    LoadListForm
End Sub

Private Sub mnuShowPrefixes_Click()
    strListSelect = "prefix"
    LoadListForm
End Sub

Private Sub mnuShowRoots_Click()
    strListSelect = "root"
    LoadListForm
End Sub

Private Sub mnuShowSuffixes_Click()
    strListSelect = "suffix"
    LoadListForm
End Sub

Private Sub optStatus_Click(Index As Integer)
    ' Enable or disable a word.
    
    RS.Edit
    
    Select Case Index
        Case 0: RS.Fields("use") = True
        Case 1: RS.Fields("use") = False
    End Select
    
    RS.Update
    
End Sub

Private Sub optWordType_Click(Index As Integer)
    ' Select word type to display.

    Select Case Index
        Case 0: strType = "prefix"
        Case 1: strType = "suffix"
        Case 2: strType = "root"
        Case 3: strType = "all"
    End Select
    
    Select Case strType
        Case "all": strSQL = "SELECT * FROM tblWords"
        Case Else:  strSQL = "SELECT * FROM tblWords WHERE type = """ & strType & """"
    End Select
    
    Set RS = DB.OpenRecordset(strSQL)
    
    intLastWord = RS.Fields("id") - 1
    intLastWordFirst = RS.Fields("id")

    i = 0
    Do
        i = i + 1
        RS.MoveNext
    Loop Until RS.EOF
    
    intLastWordMax = i
    
    cmdNext_Click

End Sub
