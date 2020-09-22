VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVacation 
   Caption         =   "Vacation Wish-List"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   480
   End
   Begin MSComctlLib.ImageList ilsTree 
      Left            =   2640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":0000
            Key             =   "ISLAND"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":0452
            Key             =   "GLOBE2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":08A4
            Key             =   "CLOSEDFLDR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVacation.frx":09FE
            Key             =   "OPENFLDR"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwVacation 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVacation.frx":0B58
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmVacation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'                   Treeview Drag & Drop Demo
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit

' The following declarations and API are used to set the form on top:
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long

' Reference to the class that will handle resizing:
Private oResize As cResize

' Scrolling up or down when dropping:
Private mintScrollDir As Integer

'------------------------------------------------------------------------
'       FORM EVENTS & INITIAL SETUP
'------------------------------------------------------------------------

Private Sub FormOnTop(bPosition As Boolean)
    Select Case bPosition
        Case True
            SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        Case False
            SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    End Select
End Sub

Private Sub Form_Load()
    ' Initialize resizing:
    Set oResize = New cResize
    oResize.InitPositions Me, False, False
    FormOnTop True
    
    ' Set the imagelist and create some folders:
    Set tvwVacation.ImageList = ilsTree
    LoadVacationTree
End Sub

Private Sub LoadVacationTree()
On Error GoTo ErrHandler

    Dim nodX                As Node
    
    ' Don't allow the tree to repaint until we are finished loading:
    TreeRedraw tvwVacation.hWnd, False
    
    With tvwVacation.Nodes
        .Clear
        ' Add the root node and make it boldface:
        .Add , , "ROOT", "Vacation Destinations", "ISLAND", "ISLAND"
        BoldTreeNode tvwVacation, tvwVacation.Nodes("ROOT")
        
        ' Add some folders:
        Set nodX = .Add("ROOT", tvwChild, "V1", "Places I Want To Visit", "CLOSEDFLDR", "OPENFLDR")
            ' Children of the above folder:
            Set nodX = .Add("V1", tvwChild, "V4", "First Choice", "CLOSEDFLDR", "OPENFLDR")
            Set nodX = .Add("V1", tvwChild, "V5", "Second Choice", "CLOSEDFLDR", "OPENFLDR")
            Set nodX = .Add("V1", tvwChild, "V6", "Third Choice", "CLOSEDFLDR", "OPENFLDR")
        Set nodX = .Add("ROOT", tvwChild, "V2", "Places I Have Already Visited", "CLOSEDFLDR", "OPENFLDR")
        Set nodX = .Add("ROOT", tvwChild, "V3", "Places To Avoid!", "CLOSEDFLDR", "OPENFLDR")
    End With
    
    ' Expand the Root but don't sort it:
    tvwVacation.Nodes("ROOT").Sorted = False
    tvwVacation.Nodes("ROOT").Expanded = True
    
    ' Allow the tree to redraw
    TreeRedraw tvwVacation.hWnd, True
    Exit Sub
    
ErrHandler:
    TreeRedraw tvwVacation.hWnd, True
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        ' User clicked the X so we just hide
        ' the form. This way our tree data
        ' will persist if the form is shown again.
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oResize = Nothing
End Sub

'------------------------------------------------------------------------
'       VACATION TREE EVENTS
'------------------------------------------------------------------------

Private Sub tvwVacation_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set the drag node:
    Set gnodDragNode = tvwVacation.HitTest(x, y)
    If Not gnodDragNode Is Nothing Then
        gnodDragNode.Tag = "tvwVacation"
    End If
End Sub

Private Sub tvwVacation_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gnodDragNode Is Nothing Then Exit Sub
    
    Dim bOkay               As Boolean
    Dim strID               As String
    
    If Button = vbLeftButton Then
        ' Test for a local drag:
        If gnodDragNode.Tag = "tvwVacation" Then
            ' No dragging the root allowed:
            If gnodDragNode.Key <> "ROOT" Then
                ' No dragging the main folders:
                strID = Left$(gnodDragNode.Key, 1)
                If strID <> "V" Then
                    bOkay = True
                    ' Display information in the label:
                    frmMain.lblMessage.Caption = MSG1 & gnodDragNode.Tag & vbNewLine & _
                             MSG2 & gnodDragNode.Text
                    ' Set the drag node and start dragging:
                    Set tvwVacation.SelectedItem = gnodDragNode
                    tvwVacation.OLEDrag
                End If
            End If
        End If
    End If
   
    If Not bOkay Then Set gnodDragNode = Nothing
    If Button = 0 Then frmMain.lblMessage.Caption = vbNullString
End Sub


Private Sub tvwVacation_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim strTemp             As String
    Dim nodTarget           As Node
    Dim nodNew              As Node
    Dim nodOldParent        As Node
    Dim strNewKey           As String
    
    ' lngKeyVal is just a hack to give us unique keys for items
    ' dragged to this tree from other trees:
    Static lngKeyVal        As Long
    
    On Error Resume Next
    
    ' Check whether the clipboard data is in our format
    strTemp = Data.GetFormat(gintClipBoardFormat)
    
    If Err Or strTemp = "False" Then
        ' it's not, so don't allow dropping
        Set gnodDragNode = Nothing
        Set tvwVacation.DropHighlight = Nothing
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    If gnodDragNode Is Nothing Then
        Set tvwVacation.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    Set nodTarget = tvwVacation.DropHighlight
    If nodTarget Is Nothing Then
        Set gnodDragNode = Nothing
        Set tvwVacation.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    ' See if we are dragging a node from the main tree to this one,
    ' or if we are just moving a node in this tree to a new location:
    If gnodDragNode.Tag = "tvwVacation" Then
        ' Just moving a local node:
        Set nodOldParent = gnodDragNode.Parent
        ' Close the old parent just for fun:
        nodOldParent.Expanded = False
        ' Now reparent to the new location:
        Set gnodDragNode.Parent = nodTarget
        Set tvwVacation.SelectedItem = gnodDragNode
        gnodDragNode.EnsureVisible
        frmMain.lblMessage.Caption = MSG4 & nodTarget.Text
    Else
        ' The node is being copied from another treeview so we
        ' construct a new key:
        lngKeyVal = lngKeyVal + 1
        strNewKey = "X" & CStr(lngKeyVal)
        Set nodNew = tvwVacation.Nodes.Add(nodTarget, tvwChild, strNewKey, gnodDragNode.Text, "GLOBE2", "GLOBE2")
        Set tvwVacation.SelectedItem = nodNew
        nodNew.EnsureVisible
        frmMain.lblMessage.Caption = MSG3
    End If
    
    Set tvwVacation.DropHighlight = Nothing
    Set gnodDragNode = Nothing
    
    Exit Sub
    
ErrHandler:
    If Not gnodDragNode Is Nothing Then Set gnodDragNode = Nothing
    
    If Err.Number = 35614 Then
        ' Circular reference error:
        MsgBox CIRCERR
    Else
        MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub tvwVacation_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    
    Dim strTemp             As String
    Dim nodTargetNode       As Node
    Dim strID               As String
    Dim strSource           As String
    Dim strKey              As String
    
    ' Start the autoscroll timer for this tree:
    If Not gnodDragNode Is Nothing Then tmrDrag.Enabled = True
    
    On Error Resume Next
    ' First check that we allow this type of data to be dropped here:
    strTemp = Data.GetFormat(gintClipBoardFormat)
    
    If Err Or strTemp = "False" Then
        ' data is not in out special format, don't allow dragging:
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    ' See what node we are over and if it is okay to drop:
    Set nodTargetNode = tvwVacation.HitTest(x, y)
    
    If nodTargetNode Is Nothing Then
        ' Not over a node so clean up and get out:
        Set tvwVacation.DropHighlight = Nothing
        Exit Sub
    End If
    
    ' Check where the drag is coming from. We only allow dragging
    ' here either from tvwMain or from this tree. We also only
    ' allow dropping onto an existing folder so we check the key:
    strSource = gnodDragNode.Tag
    strKey = Left$(nodTargetNode.Key, 1)
    
    ' No point in dropping on the source of the drag:
    If nodTargetNode.Key = gnodDragNode.Key Or strKey <> "V" Then
        Set tvwVacation.DropHighlight = Nothing
        Effect = vbDropEffectNone
    Else
        If UCase(strSource) = "TVWMAIN" Or UCase(strSource) = "TVWVACATION" Then
            ' It is okay to use this node as a target:
            Set tvwVacation.DropHighlight = nodTargetNode
        End If
    End If
    
    ' see which direction we need to scroll the treeview
    If y > 0 And y < 300 Then
        mintScrollDir = -1
    ElseIf (y < tvwVacation.Height) And y > (tvwVacation.Height - 500) Then
        mintScrollDir = 1
    Else
        mintScrollDir = 0
    End If
    
End Sub

Private Sub tvwVacation_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error Resume Next
    Dim byt()               As Byte
   
    ' From the main tree we only allow copying, not moving:
    AllowedEffects = vbDropEffectMove
    
    ' Place the key in our custom clipboard format...
    byt = gnodDragNode.Key
    ' ...and use it as the dragging data format:
    Data.SetData byt, gintClipBoardFormat

End Sub

Private Sub tvwVacation_OLECompleteDrag(Effect As Long)
    Screen.MousePointer = vbDefault
    tmrDrag.Enabled = False
End Sub

'------------------------------------------------------------------------
'       TIMER EVENT FOR AUTOEXPANDING AND AUTOSCROLLING
'------------------------------------------------------------------------

Private Sub tmrDrag_Timer()
' Used to scroll a treeview when the user drags over it
On Error GoTo ErrHandler

    Dim nodHitNode          As Node
    Static lngCount         As Long
    
    ' Don't bother if we aren't dragging:
    If gnodDragNode Is Nothing Then
        tmrDrag.Enabled = False
        Exit Sub
    End If
    
    lngCount = lngCount + 1
    
    If lngCount > 10 Then
        Set nodHitNode = tvwVacation.DropHighlight
        If nodHitNode Is Nothing Then Exit Sub
        If nodHitNode.Expanded = False Then
            nodHitNode.Expanded = True
        End If
        lngCount = 0
    End If
    If mintScrollDir <> 0 Then
        If mintScrollDir = -1 Then
            SendMessageLong tvwVacation.hWnd, WM_VSCROLL, 0, 0
        Else
            SendMessageLong tvwVacation.hWnd, WM_VSCROLL, 1, 0
        End If
    End If
 
    Exit Sub

ErrHandler:
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub


