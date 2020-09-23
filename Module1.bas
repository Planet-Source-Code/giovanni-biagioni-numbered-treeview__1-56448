Attribute VB_Name = "Module1"
'=========================================================================================
'Numbered TreeView by Giovanni Biagioni - g.biagioni@libero.it
'
'TreeView subclassing to adding a number at right of text items like
'Outlook Express
'
'The source code is based on Brad Martinez routines Many Thanks Brad!!
'Brad Martinez,  http://www.mvps.org/ccrp/
'=========================================================================================
'Usage:
'
'TreeViewNumbers_On(TreeView Object) ;start subclassing of TreeView
'TreeViewNumbers_Off ;stop subclassing of TreeView (Remember! Place in Unload)
'ItemNumber(Node, [Number], [Color]) ;setting up or retrieve number for node item
'=========================================================================================

Option Explicit

'my private
Private cltNumbers      As New Collection   'numbers collection
Private cltNumbersColor As New Collection   'number's color collection
Private lForm_hWnd      As Long             'window handle of treeview parent form
Private tTreeView       As TreeView         'treeview object
Private OldProc         As Long             'window proc old address

'messages
Private Const TV_FIRST = &H1100
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_GETITEM = (TV_FIRST + 12)
Private Const NM_CUSTOMDRAW = (-12&)
Private Const WM_NOTIFY As Long = &H4E&
Private Const CDDS_PREPAINT As Long = &H1&
Private Const CDDS_POSTPAINT As Long = &H2&
Private Const CDDS_PREERASE As Long = &H3&
Private Const CDDS_POSTERASE As Long = &H4&
Private Const CDRF_DODEFAULT = &H0
Private Const CDRF_NEWFONT = &H2
Private Const CDRF_SKIPDEFAULT = &H4
Private Const CDRF_NOTIFYITEMDRAW = &H20
Private Const CDRF_NOTIFYPOSTERASE = &H40
Private Const CDRF_NOTIFYPOSTPAINT = &H10
Private Const CDRF_NOTIFYSUBITEMDRAW = &H20
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT As Long = CDDS_ITEM Or CDDS_POSTPAINT
Private Const CDDS_ITEMPREERASE As Long = CDDS_ITEM Or CDDS_PREERASE
Private Const CDDS_ITEMPOSTERASE As Long = CDDS_ITEM Or CDDS_POSTERASE

Private Const DT_LEFT = &H0&

Private Const GWL_WNDPROC As Long = (-4&)
 
Private Type NMHDR
  hWndFrom As Long      'window handle of control sending message
  idFrom As Long        'identifier of control sending message
  code  As Long         'specifies the notification code
End Type
 
' sub struct of the NMCUSTOMDRAW struct
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
 
' generic customdraw struct
Private Type NMCUSTOMDRAW
  hdr As NMHDR
  dwDrawStage As Long
  hdc As Long
  rc As RECT
  dwItemSpec As Long
  uItemState As Long
  lItemlParam As Long
End Type
 
' treeview specific customdraw struct
Private Type NMTVCUSTOMDRAW
  nmcd As NMCUSTOMDRAW
  clrText As Long
  clrTextBk As Long
  ' if IE >= 4.0 this member of the struct can be used
  iLevel As Integer
End Type

'treeview item data
Private Type TVITEM
  mask As Long
  hItem As Long
  State As Long
  stateMask As Long
  pszText As String
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type

Private Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Function GetTVItemFromNode(hwndTV As Long, _
                                                            nod As Node) As Long
  Dim nod1 As Node
  Dim anSiblingPos() As Integer  ' contains the sibling position of the node and all it's parents
  Dim nLevel As Integer              ' hierarchical level of the node
  Dim hItem As Long
  Dim i As Integer
  Dim nPos As Integer

  Set nod1 = nod

  ' Continually work backwards from the current node to the current node's
  ' first sibling, caching the current node's sibling position in the one-based
  ' array. Then get the first sibling's parent node and start over. Keep going
  ' until the postion of the specified node's top level parent item is obtained...
  Do While (nod1 Is Nothing) = False
    nLevel = nLevel + 1
    ReDim Preserve anSiblingPos(nLevel)
    anSiblingPos(nLevel) = GetNodeSiblingPos(nod1)
    Set nod1 = nod1.Parent
  Loop

  ' Get the hItem of the first item in the treeview
  hItem = TreeView_GetRoot(hwndTV)
  If hItem Then

    ' Now work backwards through the cached node positions in the array
    ' (from the first treeview node to the specified node), obtaining the respective
    ' item handle for each node at the cached position. When we get to the
    ' specified node's position (the value of the first element in the array), we
    ' got it's hItem...
    For i = nLevel To 1 Step -1
      nPos = anSiblingPos(i)
      
      Do While nPos > 1
        hItem = TreeView_GetNextSibling(hwndTV, hItem)
        nPos = nPos - 1
      Loop
      
      If (i > 1) Then hItem = TreeView_GetChild(hwndTV, hItem)
    Next

    GetTVItemFromNode = hItem

  End If   ' hItem

End Function

Private Function GetNodeSiblingPos(nod As Node) As Integer
  
  Dim nod1 As Node
  Dim nPos As Integer
  
  Set nod1 = nod
  
  ' Keep counting up from one until the node has no more previous siblings
  Do While (nod1 Is Nothing) = False
    nPos = nPos + 1
    Set nod1 = nod1.Previous
  Loop
  
  GetNodeSiblingPos = nPos
  
End Function

Private Function TreeView_GetItem(hWnd As Long, pitem As TVITEM) As Boolean
  TreeView_GetItem = SendMessage(hWnd, TVM_GETITEM, 0, pitem)
End Function


Private Function TreeView_GetNextItem(hWnd As Long, hItem As Long, flag As Long) As Long
  TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)
End Function

Private Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
  TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)
End Function

Private Function TreeView_GetNextSibling(hWnd As Long, hItem As Long) As Long
  TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXT)
End Function

Private Function TreeView_GetPrevSibling(hWnd As Long, hItem As Long) As Long
  TreeView_GetPrevSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_PREVIOUS)
End Function

Private Function TreeView_GetRoot(hWnd As Long) As Long
  TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim hItem       As Long
Dim lNumber     As Long
Dim lOrigColor  As Long
Dim lOrigBkMode As Long
Dim rc          As RECT
Dim rcItem      As RECT
  
If iMsg = WM_NOTIFY Then
    Dim udtNMHDR As NMHDR
    CopyMemory udtNMHDR, ByVal lParam, 12&
    If udtNMHDR.code = NM_CUSTOMDRAW Then
        Dim udtNMTVCUSTOMDRAW As NMTVCUSTOMDRAW
        CopyMemory udtNMTVCUSTOMDRAW, ByVal lParam, Len(udtNMTVCUSTOMDRAW)
        Select Case udtNMTVCUSTOMDRAW.nmcd.dwDrawStage
            Case CDDS_ITEMPREPAINT
                WindowProc = CDRF_NOTIFYPOSTPAINT
                Exit Function
            Case CDDS_PREPAINT
                WindowProc = CDRF_NOTIFYITEMDRAW
                Exit Function
            Case CDDS_ITEMPOSTPAINT
                hItem = udtNMTVCUSTOMDRAW.nmcd.dwItemSpec
                lNumber = prGetItemNumber(hItem) 'retrieve number
                If lNumber <> 0 Then
                    lOrigColor = SetTextColor(udtNMTVCUSTOMDRAW.nmcd.hdc, prGetItemColor(hItem)) 'setting custom color
                    lOrigBkMode = SetBkMode(udtNMTVCUSTOMDRAW.nmcd.hdc, 2) ' opaque color
                    LSet rc = udtNMTVCUSTOMDRAW.nmcd.rc 'setting up position
                    rcItem.Left = udtNMTVCUSTOMDRAW.nmcd.dwItemSpec
                    SendMessage tTreeView.hWnd, TVM_GETITEMRECT, 1, rcItem
                    rc.Left = rc.Left + rcItem.Right + 2
                    DrawText udtNMTVCUSTOMDRAW.nmcd.hdc, "(" & CStr(lNumber) & ")", -1, rc, DT_LEFT 'draw number
                    SetTextColor udtNMTVCUSTOMDRAW.nmcd.hdc, lOrigColor 'restore colors
                    SetBkMode udtNMTVCUSTOMDRAW.nmcd.hdc, lOrigBkMode
                    WindowProc = CDRF_SKIPDEFAULT
                    Exit Function
                End If
        End Select
    End If
End If
  
WindowProc = CallWindowProc(OldProc, hWnd, iMsg, wParam, lParam)
  
End Function

Private Sub prSetItemColor(hItem As Long, lColor As Long)
On Error Resume Next
prEraseItemColor hItem
If lColor > 0 Then
    cltNumbersColor.Add lColor, CStr(hItem)
End If
End Sub


Private Sub prSetItemNumber(hItem As Long, lNumber As Long)
On Error Resume Next
prEraseItemNumber hItem
If lNumber > 0 Then
    cltNumbers.Add lNumber, CStr(hItem)
End If
End Sub


Public Function ItemNumber(nNode As Node, Optional lNumber As Variant, Optional lColor As Variant) As Long
On Error Resume Next
If IsMissing(lColor) Then lColor = &HFF0000
prSetItemColor GetTVItemFromNode(tTreeView.hWnd, nNode), CLng(lColor)
If IsMissing(lNumber) Then
    ItemNumber = prGetItemNumber(GetTVItemFromNode(tTreeView.hWnd, nNode))
Else
    prSetItemNumber GetTVItemFromNode(tTreeView.hWnd, nNode), CLng(lNumber)
End If
tTreeView.Refresh
End Function

Private Sub prEraseItemNumber(hItem As Long)
On Error Resume Next
cltNumbers.Remove (CStr(hItem))
End Sub

Private Sub prEraseItemColor(hItem As Long)
On Error Resume Next
cltNumbersColor.Remove (CStr(hItem))
End Sub





Private Function prGetItemNumber(hItem As Long) As Long
On Error Resume Next
prGetItemNumber = cltNumbers(CStr(hItem))
End Function



Private Function prGetItemColor(hItem As Long) As Long
On Error Resume Next
prGetItemColor = cltNumbersColor(CStr(hItem))
End Function



Public Sub TreeViewNumbers_On(oTreeView As TreeView)
lForm_hWnd = oTreeView.Parent.hWnd
Set tTreeView = oTreeView
OldProc = SetWindowLong(lForm_hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub TreeViewNumbers_Off()
If OldProc <> 0 Then _
    Call SetWindowLong(lForm_hWnd, GWL_WNDPROC, OldProc)
End Sub

