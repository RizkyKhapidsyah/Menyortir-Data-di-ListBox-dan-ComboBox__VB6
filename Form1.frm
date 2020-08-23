VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menyortir Data di ListBox dan ComboBox"
   ClientHeight    =   1740
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ReSort(L As Control)   'Fungsi untuk menyortir data
  Dim P%, PP%, C%, Pre$, S$, V&, NewPos%, CheckIt%
  Dim TempL$, TempItemData&, S1$
  
  For P = 0 To L.ListCount - 1
    S = L.List(P)
    For C = 1 To Len(S)
        V = Val(Mid$(S, C))
        If V > 0 Then Exit For
    Next
    If V > 0 Then
        If C > 1 Then Pre = Left$(S, C - 1)
        NewPos = -1
        For PP = P + 1 To L.ListCount - 1
            CheckIt = False
            S1 = L.List(PP)
            If Pre <> "" Then
               If InStr(S1, Pre) = 1 Then CheckIt = _
                  True
            Else
                If Val(S1) > 0 Then CheckIt = True
            End If
            If CheckIt Then
               If Val(Mid$(S1, C)) < V Then NewPos = PP
            Else
                Exit For
            End If
        Next
        If NewPos > -1 Then
            TempL = L.List(P)
            TempItemData = L.ItemData(P)
            L.RemoveItem (P)
            L.AddItem TempL, NewPos
            L.ItemData(L.NewIndex) = TempItemData
            P = P - 1
        End If
    End If
  Next
  Exit Sub
End Sub

Private Sub Command1_Click()
   Call ReSort(List1)  'Sortir data di listbox
End Sub
Private Sub Command2_Click()
   Call ReSort(Combo1)  'Sortir data di combobox
End Sub

Private Sub Form_Load()
    'Tambahkan item data ke dalam listbox
    List1.AddItem "File3.gif"
    List1.AddItem "File2.gif"
    List1.AddItem "File10.gif"
    List1.AddItem "File1.gif"
    'Tambahkan item data ke dalam combobox
    Combo1.AddItem "File3.gif"
    Combo1.AddItem "File2.gif"
    Combo1.AddItem "File10.gif"
    Combo1.AddItem "File1.gif"
End Sub


