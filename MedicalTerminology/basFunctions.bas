Attribute VB_Name = "basFunctions"
 
 ' basFunctions
Option Explicit

Public Function RandomRange(Low As Integer, High As Integer) As Integer

    Randomize
    RandomRange = Int((High - Low + 1) * Rnd + Low)
    
End Function

Public Sub LoadListForm()

    Load frmList
    frmList.Show

End Sub
