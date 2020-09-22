Attribute VB_Name = "modGeneral"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'                   Treeview Drag & Drop Demo
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit

'------------------------------------
' General constants used by this demo
'------------------------------------
Public Const INTRO As String = "Treeview Drag-n-Drop Demo Guide:" & vbNewLine & vbNewLine & _
    "In this demo I tried to show several treeview methods for drag and drop, sorting, copying " & _
    "nodes between forms, and changing images based on the number of children." & _
    vbNewLine & vbNewLine & "The nodes in the first tree with " & _
    "the list of countries can be copied to the other trees by dragging a node, but the tree " & _
    "itself cannot be modified. You can drag items to the tree with the seven continents, but you " & _
    "can only drop in a folder. No mixing countries! You can also move items within that tree to " & _
    "a new location." & vbNewLine & vbNewLine & "The third tree creates a new key when an item is dropped there, so anything " & _
    "goes! You can add items more than once or mix countries. If an item has children the image " & _
    "is a folder; otherwise the image is a globe. Timers provide both auto-scrolling and auto-expanding " & _
    "of the tree that the cursor is over." & vbNewLine & vbNewLine & "You can also learn how to drag an item to a different " & _
    "form in a project. From the menu select 'Show Vacation Form' and drag an item to one of the " & _
    "folders. Like the other trees, you can also drag a node to a new location within the same tree. " & _
    "Instead of relying on the SetData OLE method, this project will also demonstrate registering " & _
    "a custom format for the drag data. This technique gives us much more control over what we allow " & _
    "to be dropped. Credit for this method goes to Chris Eastwood. The method of getting a collection " & _
    "of countries from the registry was written by Dinesh Asanka and posted at www.vb-helper.com"
 
Public Const MSG1 As String = "Origin: "
Public Const MSG2 As String = "Item: "
Public Const MSG3 As String = "Drop successful."
Public Const MSG4 As String = "Move successful. The new parent of this item is "
Public Const MSG5 As String = "Action cancelled."
Public Const KEYERR As String = "That item already exists in the tree. The drop was cancelled."
Public Const CIRCERR As String = "Dropping a parent on a child creates a circular reference. The drop was cancelled."

'------------------------------------
' API Constants, Types, & Declares
'------------------------------------
' Treeview messages and styles
Public Const TV_FIRST           As Long = &H1100
Public Const TVM_GETNEXTITEM    As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM        As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM        As Long = (TV_FIRST + 13)
Public Const TVIF_STATE         As Long = &H8
Public Const TVIS_BOLD          As Long = &H10
Public Const TVGN_CARET         As Long = &H9
Public Const WM_SETREDRAW       As Long = &HB
Public Const WM_VSCROLL         As Long = &H115

' Treeview Item Structure
Public Type TVITEM
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

' API declare for registering a custom format
Public Declare Function RegisterClipboardFormat Lib "user32" _
        Alias "RegisterClipboardFormatA" _
        (ByVal lpString As String) As Integer

' Used for turning the treeview drawing off when loading items
Public Declare Function SendMessageLong Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

' Need the 'Any' parameter for the BoldTreeNode sub
Public Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long


' ***** These are declared global so we can drag between forms *****
' Registering a clipboard format returns an integer:
Public gintClipBoardFormat As Integer

' Declare a node variable we use when dragging:
Public gnodDragNode As Node

' We pass a string to register a clipboard format:
Public Const CLIPBOARD_NAME As String = "TreeviewDemo"

'------------------------------------
' Public Methods Used by Both Forms
'------------------------------------
Public Sub TreeRedraw(ByVal lHwnd As Long, ByVal bRedraw As Boolean)
' Turn treeview redraw on/off
    SendMessageLong lHwnd, WM_SETREDRAW, bRedraw, 0
End Sub

Public Sub BoldTreeNode(tvw As TreeView, nNode As Node)
' Routine from vbnet
On Error GoTo ErrHandler

    Dim TVI As TVITEM
    Dim lRet As Long
    Dim hItemTV As Long
    Dim lHwnd As Long
    
    Set tvw.SelectedItem = nNode
    
    lHwnd = tvw.hWnd
    hItemTV = SendMessageLong(lHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
    
    If hItemTV > 0 Then
        With TVI
            .hItem = hItemTV
            .mask = TVIF_STATE
            .stateMask = TVIS_BOLD
            lRet = SendMessage(lHwnd, TVM_GETITEM, 0&, TVI)
            .State = TVIS_BOLD
        End With
        lRet = SendMessage(lHwnd, TVM_SETITEM, 0&, TVI)
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub


