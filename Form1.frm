VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Numbered Treeview  Example"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Numbers Color"
      Height          =   1935
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      Begin VB.OptionButton Option4 
         Caption         =   "Random"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Green"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Red"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blue"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin ComctlLib.TreeView TreeView2 
      Height          =   3555
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6271
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
      
Dim i       As Integer
Dim j       As Integer
Dim k       As Integer
Dim Node1   As Node
Dim Node2   As Node
Dim Node3   As Node

'populate treview with demo item
With TreeView2
    .HideSelection = False
    .Indentation = 19 * Screen.TwipsPerPixelX
    .LineStyle = tvwRootLines
    
    ' number of Nodes at each respective hierarchical level
    Const nNodes1 = 2
    Const nNodes2 = 3
    Const nNodes3 = 4
    
    ' Fill up the treeview...
    For i = 1 To nNodes1
        Set Node1 = .Nodes.Add(, , , "Root " & i)
        For j = 1 To nNodes2
            Set Node2 = .Nodes.Add(Node1.Index, tvwChild, , "Root " & i & " Child " & j)
            For k = 1 To nNodes3
                Set Node3 = .Nodes.Add(Node2.Index, tvwChild, , _
                " GrandChild " & (nNodes2 * nNodes3 * (i - 1)) + (nNodes3 * (j - 1)) + k)
            Next k
        Next j
    Next i
    
    'setting up treeview numbers functions
    'i use subclassing REMENBER!! place TreeViewNumbers_Off on unload event
    TreeViewNumbers_On TreeView2
    
    'setting number usesing children numbers
    For i = 1 To .Nodes.Count
        ItemNumber .Nodes(i), .Nodes(i).Children
    Next
    
End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 TreeViewNumbers_Off
End Sub

Private Sub Option1_Click()

Dim i As Integer

'setting ItemNumber to Blue color
For i = 1 To TreeView2.Nodes.Count
  ItemNumber TreeView2.Nodes(i), , vbBlue
Next

End Sub

Private Sub Option2_Click()

Dim i As Integer

'setting ItemNumber to Red color
For i = 1 To TreeView2.Nodes.Count
  ItemNumber TreeView2.Nodes(i), , vbRed
Next

End Sub


Private Sub Option3_Click()

Dim i As Integer

'setting ItemNumber to Green color
For i = 1 To TreeView2.Nodes.Count
  ItemNumber TreeView2.Nodes(i), , vbGreen
Next

End Sub


Private Sub Option4_Click()

Dim i As Integer
Dim c As Long

'setting ItemNumber to Random color
For i = 1 To TreeView2.Nodes.Count
  Select Case Int((3 - 1 + 1) * Rnd + 1)
  Case 1: c = vbRed
  Case 2: c = vbGreen
  Case 3: c = vbBlue
  End Select
  ItemNumber TreeView2.Nodes(i), , c
Next

End Sub


