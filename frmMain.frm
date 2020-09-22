VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Treeview Drag and Drop Demo"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   480
      Top             =   6120
   End
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   360
      Top             =   6000
   End
   Begin MSComctlLib.ImageList ilsTree 
      Left            =   0
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "ROOTCLOSED"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27B2
            Key             =   "ROOTOPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F64
            Key             =   "CLOSEDFLDR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50BE
            Key             =   "OPENFLDR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5218
            Key             =   "GLOBE1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B4B2
            Key             =   "GLOBE2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B904
            Key             =   "MAP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BD56
            Key             =   "SEARCHGLOBE"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BEB0
            Key             =   "TARGETFLDR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1CA
            Key             =   "PEOPLEFLDR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C4E4
            Key             =   "MAN"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C936
            Key             =   "ISLAND"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CD88
            Key             =   "POINTER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   177
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwContinent 
      Height          =   5295
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.TreeView tvwGroups 
      Height          =   5295
      Left            =   6000
      TabIndex        =   2
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Groups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "This treeview allows duplicates. You can drop an item at any level."
      Top             =   600
      Width           =   2685
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Continents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Test your geography knowledge! Drag countries to the folder where you think they belong."
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Countries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "A list of all countries. You can copy from this list by dragging, but you can't remove or add items!"
      Top             =   600
      Width           =   2640
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuFormsTop 
      Caption         =   "&Other Forms"
      Begin VB.Menu mnuForms 
         Caption         =   "Show &Intro Form"
         Index           =   0
      End
      Begin VB.Menu mnuForms 
         Caption         =   "Show &Vacation Form"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'                   Treeview Drag & Drop Demo
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'------------------------------------------------------------------------
' Programmer:   Treeview coding by Brent Culpepper (IDontKnow)
'------------------------------------------------------------------------
'
' Credits:      Full credit goes to Dinesh Asanka for the code to
'               retrieve the country list from the registry. This can
'               be found posted at:
'               http://www.vb-helper.com/howto_list_countries.html
'
'               The auto-scroll was originally posted at MSDN but was
'               modified and improved by Chris Eastwood in an article at:
'               http://www.codeguru.com/
'               You can find more of Chris Eastwood's work at:
'               http://www.vbcodelibrary.co.uk/index.php
'
'               The code to bold a node was posted at www.vbnet.mvps.org
'
'               I created the resizing class module by modifying an
'               ActiveX control written by Francesco Balena and posted at:
'               http://www.vb2themax.com
'
'               I offer my sincere gratitude to all the above for their
'               willingness to share their knowledge with others!
'
'------------------------------------------------------------------------
' Date:         April 15, 2004
'------------------------------------------------------------------------
'
' Purpose:      1. Demonstrates dragging and dropping between treeviews.
'               2. Demonstates using a custom data format to ensure
'                  that the data dropped is in a valid format.
'               3. Demonstrates using a timer to expand and scroll a node.
'               4. Shows how to change the image if a node has children.
'               5. Demonstrates dragging between forms.
'               6. Shows how to use Effects to either copy, move, or
'                  cancel an OLEDrag operation.
'               7. Demonstrates dragging a node to a new parent in the tree.
'               8. Shows how to handle common drag/drop errors.
'               9. Shows how to validate the dragged item and only allow
'                  dropping if it meets certain conditions.
'              10. As a bonus, shows how to use a resizing class module
'                  to handle basic resizing operations!
'
'------------------------------------------------------------------------
' Requires:     Microsoft Windows Common Controls 6.0 (MSCOMCTL.OCX)
'------------------------------------------------------------------------
Option Explicit


' Declarations required for retrieving the country list:
Private Const ValueName As String = "Name"
Private Const MasterKey As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Country List\"

' Reference to the class that will handle resizing:
Private oResize As cResize
'------------------------------------------------------------------------
' The following are the declarations we need for the treeview operations

Private Const CONT1 As String = "Africa"
Private Const CONT2 As String = "Antarctica"
Private Const CONT3 As String = "Asia"
Private Const CONT4 As String = "Australia"
Private Const CONT5 As String = "Europe"
Private Const CONT6 As String = "North America"
Private Const CONT7 As String = "South America"


Private Const GROUP1 As String = "United Nations Members"
Private Const GROUP2 As String = "NATO Members"
Private Const GROUP3 As String = "High Cost of Living"
Private Const GROUP4 As String = "Don't Drink the Water!"


' Scrolling up or down when dropping:
Private mintScrollDir As Integer

'------------------------------------------------------------------------
'       FORM EVENTS & INITIAL SETUP
'------------------------------------------------------------------------

Private Sub Form_Load()
On Error GoTo ErrHandler

    ' Initialize resizing:
    Set oResize = New cResize
    oResize.InitPositions Me, False, False
    
    ' Register our format for drag/drop operations:
    gintClipBoardFormat = RegisterClipboardFormat(CLIPBOARD_NAME)
    
    ' Assign our imagelist to the trees:
    Set tvwMain.ImageList = ilsTree
    Set tvwContinent.ImageList = ilsTree
    Set tvwGroups.ImageList = ilsTree
    
    ' Load the main tree with the country list:
    LoadMainTreeWithCountries
    
    ' Load the second tree with folders for continents:
    LoadContinentsTree
    
    ' Load some folders in the Group Tree which will
    ' allow duplicate items:
    LoadGroupTree
        
    Exit Sub
    
ErrHandler:
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Clean up:
    Set oResize = Nothing
    If Not gnodDragNode Is Nothing Then Set gnodDragNode = Nothing
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.name <> Me.name Then
            Unload frm
            Set frm = Nothing
        End If
    Next
    Unload Me
End Sub

Private Sub LoadMainTreeWithCountries()
On Error GoTo ErrHandler
    
    Dim vItem               As Variant
    Dim lCount              As Long
    Dim strKey              As String
    Dim nodX                As Node
    Dim CountryCol          As Collection

    
    ' Get a list of countries to use as tree data:
    If CheckRegistryKey(HKEY_LOCAL_MACHINE, MasterKey) Then
        Dim KeyCol As Collection
        Dim TheKey As Variant
        Set KeyCol = EnumRegistryKeys(HKEY_LOCAL_MACHINE, MasterKey)
        Set CountryCol = New Collection
        lCount = 1
        For Each TheKey In KeyCol
             If TheKey <> "800" And GetRegistryValue(HKEY_LOCAL_MACHINE, MasterKey & TheKey, "InternationalRule", "") <> "00EFG#" Then
                CountryCol.Add GetRegistryValue(HKEY_LOCAL_MACHINE, MasterKey & TheKey, ValueName, ""), CStr(lCount)
             End If
             lCount = lCount + 1
        Next
    End If
    
    ' We can just reuse the lCount variable for our string-key:
    lCount = 1

    ' Don't allow the tree to repaint until we are finished loading:
    TreeRedraw tvwMain.hWnd, False
    
    With tvwMain.Nodes
        .Clear
        ' Add the root node and make it boldface
        .Add , , "ROOT", "All Countries"
        BoldTreeNode tvwMain, tvwMain.Nodes("ROOT")
        
        ' Load the countries as a child of the root. The string-key
        ' MUST start with a letter instead of a number:
        For Each vItem In CountryCol
            Set nodX = .Add("ROOT", tvwChild, "C" & CStr(lCount), vItem)
            lCount = lCount + 1
        Next
    End With
    
    ' Set the closed and open images for the root node:
    tvwMain.Nodes("ROOT").Image = "ROOTCLOSED"
    tvwMain.Nodes("ROOT").ExpandedImage = "ROOTOPEN"
    
    ' Expand the Root and sort the countries:
    tvwMain.Nodes("ROOT").Sorted = True
    tvwMain.Nodes("ROOT").Expanded = True
    
    ' Allow the tree to redraw
    TreeRedraw tvwMain.hWnd, True

    Exit Sub
  
ErrHandler:
    TreeRedraw tvwMain.hWnd, True
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub LoadContinentsTree()
On Error GoTo ErrHandler

    Dim nodX                As Node
    
    ' Don't allow the tree to repaint until we are finished loading:
    TreeRedraw tvwContinent.hWnd, False
    
    With tvwContinent.Nodes
        .Clear
        ' Add the root node and make it boldface:
        .Add , , "ROOT", "Seven Continents", "GLOBE1"
        BoldTreeNode tvwContinent, tvwContinent.Nodes("ROOT")
        
        ' Add the folders. We won't allow dragging folders,
        ' but the contents can be moved to a new location.
        ' Folders will be indicated by "F" in the key:
        Set nodX = .Add("ROOT", tvwChild, "F1", CONT1)
        Set nodX = .Add("ROOT", tvwChild, "F2", CONT2)
        Set nodX = .Add("ROOT", tvwChild, "F3", CONT3)
        Set nodX = .Add("ROOT", tvwChild, "F4", CONT4)
        Set nodX = .Add("ROOT", tvwChild, "F5", CONT5)
        Set nodX = .Add("ROOT", tvwChild, "F6", CONT6)
        Set nodX = .Add("ROOT", tvwChild, "F7", CONT7)
    End With
    
    ' Set the image and expanded image:
    ' (We could also have done this when we created the nodes)
    For Each nodX In tvwContinent.Nodes
        If nodX.Key <> "ROOT" Then
            nodX.Image = "CLOSEDFLDR"
            nodX.ExpandedImage = "OPENFLDR"
        End If
    Next
    
    ' Expand the Root and sort the list:
    tvwContinent.Nodes("ROOT").Sorted = True
    tvwContinent.Nodes("ROOT").Expanded = True
    
    ' Allow the tree to redraw
    TreeRedraw tvwContinent.hWnd, True
    Exit Sub
    
ErrHandler:
    TreeRedraw tvwContinent.hWnd, True
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub LoadGroupTree()
On Error GoTo ErrHandler

    Dim nodX                As Node
    
    ' Don't allow the tree to repaint until we are finished loading:
    TreeRedraw tvwGroups.hWnd, False
    
    With tvwGroups.Nodes
        .Clear
        ' Add the root node and make it boldface:
        .Add , , "ROOT", "Groups of Countries"
        BoldTreeNode tvwGroups, tvwGroups.Nodes("ROOT")
        
        ' Add some folders. This tree allows duplicates because
        ' some countries will belong in more than one folder:
        Set nodX = .Add("ROOT", tvwChild, "I1", GROUP1)
        Set nodX = .Add("ROOT", tvwChild, "I2", GROUP2)
        Set nodX = .Add("ROOT", tvwChild, "I3", GROUP3)
        Set nodX = .Add("ROOT", tvwChild, "I4", GROUP4)
    End With
    
    ' Note: When items are added to this tree, the node's image
    ' will depend on whether it has children or not. Nodes with
    ' children will have folders and those without have a globe.
    ' This tree will also allow you to add a country to another
    ' country.
    ' Set the images for the initial nodes:
    For Each nodX In tvwGroups.Nodes
        If nodX.Key <> "ROOT" Then
            nodX.Image = "GLOBE2"
            nodX.ExpandedImage = "GLOBE2"
        Else
            nodX.Image = "ROOTCLOSED"
            nodX.ExpandedImage = "ROOTOPEN"
        End If
    Next
    
    ' This tree is not expanded so you can see how
    ' the auto-scroll and auto-expand works. Here
    ' we just set the sorted property to True:
    tvwGroups.Nodes("ROOT").Sorted = True
    tvwGroups.Nodes("ROOT").Expanded = False
    
    ' Allow the tree to redraw
    TreeRedraw tvwGroups.hWnd, True
    Exit Sub
    
ErrHandler:
    TreeRedraw tvwGroups.hWnd, True
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
End Sub


'------------------------------------------------------------------------
'       EVENTS FOR THE MAIN TREE (COUNTRY LIST)
'------------------------------------------------------------------------
Private Sub tvwMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set the drag node:
    Set gnodDragNode = tvwMain.HitTest(x, y)
    If Not gnodDragNode Is Nothing Then
        ' Use the tag property to indicate the source of the drag:
        gnodDragNode.Tag = "tvwMain"
        ' Display current action in the label:
        lblMessage.Caption = MSG1 & gnodDragNode.Tag & vbNewLine & _
                             MSG2 & gnodDragNode.Text
    End If
End Sub

Private Sub tvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gnodDragNode Is Nothing Then Exit Sub
    
    ' This flag is for a node that is okay to drag:
    Dim bOkay               As Boolean
    
    If Button = vbLeftButton Then
        If gnodDragNode.Key <> "ROOT" Then
            ' Start Dragging
            bOkay = True
            Set tvwMain.SelectedItem = gnodDragNode
            tvwMain.OLEDrag
        End If
    End If
    If Not bOkay Then Set gnodDragNode = Nothing
End Sub

Private Sub tvwMain_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error Resume Next
    Dim byt()               As Byte
   
    ' From the main tree we only allow copying, not moving:
    AllowedEffects = vbDropEffectCopy
    
    ' Place the key in our custom clipboard format...
    byt = gnodDragNode.Key
    ' ...and use it as the dragging data format:
    Data.SetData byt, gintClipBoardFormat
    
End Sub

'------------------------------------------------------------------------
'       EVENTS FOR THE CONTINENT TREE
'------------------------------------------------------------------------
Private Sub tvwContinent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set the drag node:
    Set gnodDragNode = tvwContinent.HitTest(x, y)
    If Not gnodDragNode Is Nothing Then
        ' Use the drag node's Tag property to track the origin tree:
        gnodDragNode.Tag = "tvwContinent"
    End If
End Sub

Private Sub tvwContinent_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gnodDragNode Is Nothing Then Exit Sub
    
    ' This flag is for a node that is okay to drag:
    Dim bOkay               As Boolean
    
    ' Because we have several trees, we need to keep track of both the
    ' source and the tree we are over. This way we can tell if we are
    ' starting a drag on this tree, passing by, or waiting to drop.
    
    If Button = vbLeftButton Then
        ' Test for a local drag:
        If gnodDragNode.Tag = "tvwContinent" Then
            ' No dragging the root allowed:
            If gnodDragNode.Key <> "ROOT" Then
                ' No dragging the continent folders:
                Dim strID As String
                strID = Left$(gnodDragNode.Key, 1)
                If strID <> "F" Then
                    bOkay = True
                    ' Display information in the label:
                    lblMessage.Caption = MSG1 & gnodDragNode.Tag & vbNewLine & _
                             MSG2 & gnodDragNode.Text
                    ' Set the drag node and start dragging:
                    Set tvwContinent.SelectedItem = gnodDragNode
                    tvwContinent.OLEDrag
                End If
            End If
        End If
    End If
    If Not bOkay Then Set gnodDragNode = Nothing
    ' This just clears the label if a drop is cancelled:
    If Button = 0 Then lblMessage.Caption = vbNullString
End Sub

Private Sub tvwContinent_OLECompleteDrag(Effect As Long)
    Screen.MousePointer = vbDefault
    tmrDrag(0).Enabled = False
End Sub

Private Sub tvwContinent_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim strTemp             As String
    Dim nodTarget           As Node
    Dim nodNew              As Node
    Dim nodOldParent        As Node
    
    On Error Resume Next
    
    ' Check whether the clipboard data is in our format
    strTemp = Data.GetFormat(gintClipBoardFormat)
    
    If Err Or strTemp = "False" Then
        ' it's not, so don't allow dropping
        Set gnodDragNode = Nothing
        Set tvwContinent.DropHighlight = Nothing
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    If gnodDragNode Is Nothing Then
        Set tvwContinent.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    Set nodTarget = tvwContinent.DropHighlight
    If nodTarget Is Nothing Then
        Set gnodDragNode = Nothing
        Set tvwContinent.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    ' See if we are dragging a node from the other tree to this one,
    ' or if we are just moving a node in this tree to a new location:
    If gnodDragNode.Tag = "tvwContinent" Then
        Set nodOldParent = gnodDragNode.Parent
        nodOldParent.Expanded = False
        ' Just moving a local node so change the parent:
        Set gnodDragNode.Parent = nodTarget
        gnodDragNode.Sorted = True
        Set tvwContinent.SelectedItem = gnodDragNode
        gnodDragNode.EnsureVisible
        lblMessage.Caption = MSG4 & nodTarget.Text
    Else
        ' The node is being copied from the another treeview:
        Set nodNew = tvwContinent.Nodes.Add(nodTarget, tvwChild, gnodDragNode.Key, gnodDragNode.Text, "GLOBE2")
        nodNew.Sorted = True
        Set tvwContinent.SelectedItem = nodNew
        nodNew.EnsureVisible
        lblMessage.Caption = MSG3
    End If
    
    Set tvwContinent.DropHighlight = Nothing
    Set gnodDragNode = Nothing
    
    Exit Sub
    
ErrHandler:
    ' Check for duplicate key error. This happens if you try to add
    ' a node that already exists:
    If Err.Number = 35602 Then
        MsgBox KEYERR
    Else
        MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub tvwContinent_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    
    Dim strTemp             As String
    Dim nodTargetNode       As Node
    Dim strID               As String
                
    ' Start the autoscroll timer for this tree:
    If Not gnodDragNode Is Nothing Then
        tmrDrag(0).Enabled = True
        If gnodDragNode.Tag = "tvwGroups" Or _
        gnodDragNode.Tag = "tvwVacation" Then
            ' Don't allow dropping from the group tree because
            ' it uses a different key. We don't want duplicates.
            ' Also don't bother with the vacation tree on the
            ' other form:
            Set tvwContinent.DropHighlight = Nothing
            Effect = vbDropEffectNone
            Exit Sub
        End If
    End If
    
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
    Set nodTargetNode = tvwContinent.HitTest(x, y)
    
    If nodTargetNode Is Nothing Then
        ' Not over a node so clean up and get out:
        Set tvwContinent.DropHighlight = Nothing
        Exit Sub
    End If
    
    ' See if we are dropping on a folder, or trying to drop
    ' a country on another country (that could start a war!)
    ' The left character of our countries is "C" and the left
    ' of the folders is a "F". We also don't allow dropping
    ' on the root item:
    strID = Left$(nodTargetNode.Key, 1)
                
    If nodTargetNode.Key = gnodDragNode.Key Or _
    nodTargetNode.Key = "ROOT" Or strID <> "F" Then
        Set tvwContinent.DropHighlight = Nothing
        Effect = vbDropEffectNone
    Else
        ' It is okay to use this node as a target:
        Set tvwContinent.DropHighlight = nodTargetNode
    End If
    
    ' see which direction we will need to scroll the treeview:
    If y > 0 And y < 300 Then
        mintScrollDir = -1
    ElseIf (y < tvwContinent.Height) And y > (tvwContinent.Height - 500) Then
        mintScrollDir = 1
    Else
        mintScrollDir = 0
    End If
    
End Sub

Private Sub tvwContinent_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error Resume Next
    Dim byt()               As Byte
   
    ' From the main tree we only allow copying, not moving:
    AllowedEffects = vbDropEffectMove
    
    ' Place the key in our custom clipboard format...
    byt = gnodDragNode.Key
    ' ...and use it as the dragging data format:
    Data.SetData byt, gintClipBoardFormat
   
End Sub

'------------------------------------------------------------------------
'       EVENTS FOR THE GROUPS TREE
'------------------------------------------------------------------------

Private Sub tvwGroups_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set the drag node:
    Set gnodDragNode = tvwGroups.HitTest(x, y)
    If Not gnodDragNode Is Nothing Then
        gnodDragNode.Tag = "tvwGroups"
    End If
End Sub

Private Sub tvwGroups_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gnodDragNode Is Nothing Then Exit Sub
    
    Dim bOkay               As Boolean
    
    ' Because we have several trees, we need to keep track of both the
    ' source and the tree we are over. This way we can tell if we are
    ' starting a drag on this tree, passing by, or waiting to drop.
    
    If Button = vbLeftButton Then
        If gnodDragNode.Tag = "tvwGroups" Then
            If gnodDragNode.Key <> "ROOT" Then
                bOkay = True
                ' Display current action in the label:
                lblMessage.Caption = MSG1 & gnodDragNode.Tag & vbNewLine & _
                        MSG2 & gnodDragNode.Text
                ' Start Dragging
                Set tvwGroups.SelectedItem = gnodDragNode
                tvwGroups.OLEDrag
            End If
        End If
    End If
    If Not bOkay Then Set gnodDragNode = Nothing
    If Button = 0 Then lblMessage.Caption = vbNullString
End Sub

Private Sub tvwGroups_OLECompleteDrag(Effect As Long)
    Screen.MousePointer = vbDefault
    tmrDrag(1).Enabled = False
End Sub

Private Sub tvwGroups_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim strTemp             As String
    Dim nodTarget           As Node
    Dim nodNew              As Node
    Dim nodOldParent        As Node
    Dim strNewKey           As String
    ' lngKeyVal is just a hack to give us unique keys for items
    ' dragged to this tree from other trees:
    Static lngKeyVal        As Long
    
    ' This is just to account for the 4 existing items when program loads:
    If lngKeyVal < 4 Then lngKeyVal = 4
    
    On Error Resume Next
    
    ' Check whether the clipboard data is in our format
    strTemp = Data.GetFormat(gintClipBoardFormat)
    
    If Err Or strTemp = "False" Then
        ' it's not, so don't allow dropping
        Set gnodDragNode = Nothing
        Set tvwGroups.DropHighlight = Nothing
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    If gnodDragNode Is Nothing Then
        Set tvwGroups.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    Set nodTarget = tvwGroups.DropHighlight
    If nodTarget Is Nothing Then
        Set gnodDragNode = Nothing
        Set tvwGroups.DropHighlight = Nothing
        Effect = vbDropEffectNone
        Exit Sub
    End If
    
    ' See if we are dragging a node from the other tree to this one,
    ' or if we are just moving a node in this tree to a new location:
    If gnodDragNode.Tag = "tvwGroups" Then
        ' Just moving a local node:
        ' First get a reference to the old parent:
        Set nodOldParent = gnodDragNode.Parent
        ' See if we need to change the image in case the old parent
        ' now has no children:
        If nodOldParent.Children <= 1 Then
            If nodOldParent.Key <> "ROOT" Then
                nodOldParent.ExpandedImage = "GLOBE2"
                nodOldParent.Image = "GLOBE2"
                nodOldParent.Expanded = False
            End If
        Else
            nodOldParent.Expanded = False
        End If
        ' Now reparent to the new location:
        Set gnodDragNode.Parent = nodTarget
        Set tvwGroups.SelectedItem = gnodDragNode
        gnodDragNode.EnsureVisible
        lblMessage.Caption = MSG4 & nodTarget.Text
    Else
        ' The node is being copied from the another treeview so we
        ' construct a new key:
        lngKeyVal = lngKeyVal + 1
        strNewKey = "I" & CStr(lngKeyVal)
        Set nodNew = tvwGroups.Nodes.Add(nodTarget, tvwChild, strNewKey, gnodDragNode.Text, "GLOBE2")
        nodNew.Sorted = True
        Set tvwGroups.SelectedItem = nodNew
        nodNew.EnsureVisible
        lblMessage.Caption = MSG3
    End If
    
    ' Set the image of the parent to folder view:
    nodTarget.Image = "CLOSEDFLDR"
    nodTarget.ExpandedImage = "OPENFLDR"
    nodTarget.Sorted = True
    Set tvwGroups.DropHighlight = Nothing
    Set gnodDragNode = Nothing
    
    Exit Sub
    
ErrHandler:
    If Not gnodDragNode Is Nothing Then Set gnodDragNode = Nothing
    ' A circular reference error happens if you try and drop a parent
    ' node onto one of its child nodes. This is not good!
    If Err.Number = 35614 Then
        MsgBox CIRCERR
    Else
        MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub tvwGroups_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    
    Dim strTemp             As String
    Dim nodTargetNode       As Node
    Dim strID               As String
                
    ' Start the autoscroll timer for this tree:
    If Not gnodDragNode Is Nothing Then tmrDrag(1).Enabled = True
                
    On Error Resume Next
    ' First check that we allow this type of data to be dropped here
    strTemp = Data.GetFormat(gintClipBoardFormat)
    
    If Err Or strTemp = "False" Then
        ' data is not in out special format, don't allow dropping!
        Err.Clear
        Effect = vbDropEffectNone
        Exit Sub
    End If
        
    ' See what node we are over and if it is okay to drop:
    Set nodTargetNode = tvwGroups.HitTest(x, y)
    
    If nodTargetNode Is Nothing Then
        ' Not over a node so clean up and get out:
        Set tvwGroups.DropHighlight = Nothing
        Exit Sub
    End If
              
    If nodTargetNode.Key = gnodDragNode.Key Then
        Set tvwGroups.DropHighlight = Nothing
        Effect = vbDropEffectNone
    Else
        ' It is okay to use this node as a target:
        Set tvwGroups.DropHighlight = nodTargetNode
    End If
    
    ' see which direction we need to scroll the treeview
    If y > 0 And y < 300 Then
        mintScrollDir = -1
    ElseIf (y < tvwGroups.Height) And y > (tvwGroups.Height - 500) Then
        mintScrollDir = 1
    Else
        mintScrollDir = 0
    End If
    
    
End Sub

Private Sub tvwGroups_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error Resume Next
    Dim byt()               As Byte
   
    ' With local drags we use the effect "move" instead of "copy":
    AllowedEffects = vbDropEffectMove
    
    ' Place the key in our custom clipboard format...
    byt = gnodDragNode.Key
    ' ...and use it as the dragging data format:
    Data.SetData byt, gintClipBoardFormat

End Sub

'------------------------------------------------------------------------
'       TIMER EVENT FOR AUTOEXPANDING AND AUTOSCROLLING
'------------------------------------------------------------------------

Private Sub tmrDrag_Timer(index As Integer)
' Used to scroll a treeview when the user drags over it. Originally from
' MSDN but modified by Chris Eastwood.

On Error GoTo ErrHandler

    Dim nodHitNode          As Node
    Static lngCount         As Long
    
    ' Don't bother if we aren't dragging:
    If gnodDragNode Is Nothing Then
        tmrDrag(index).Enabled = False
        Exit Sub
    End If
    
    lngCount = lngCount + 1
    
    Select Case index
        Case 0
            
            If lngCount > 10 Then
                Set nodHitNode = tvwContinent.DropHighlight
                If nodHitNode Is Nothing Then Exit Sub
                If nodHitNode.Expanded = False Then
                    If nodHitNode.Children Then
                        nodHitNode.Expanded = True
                    End If
                End If
                lngCount = 0
            End If
            If mintScrollDir <> 0 Then
                If mintScrollDir = -1 Then
                    SendMessageLong tvwContinent.hWnd, WM_VSCROLL, 0, 0
                Else
                    SendMessageLong tvwContinent.hWnd, WM_VSCROLL, 1, 0
                End If
            End If
        Case 1
            If lngCount > 10 Then
                Set nodHitNode = tvwGroups.DropHighlight
                If nodHitNode Is Nothing Then Exit Sub
                 If nodHitNode.Expanded = False Then
                    If nodHitNode.Children Then
                        nodHitNode.Expanded = True
                    End If
                End If
                lngCount = 0
            End If
            If mintScrollDir <> 0 Then
                If mintScrollDir = -1 Then
                    SendMessageLong tvwGroups.hWnd, WM_VSCROLL, 0, 0
                Else
                    SendMessageLong tvwGroups.hWnd, WM_VSCROLL, 1, 0
                End If
            End If
    End Select
    
    Exit Sub

ErrHandler:
    MsgBox "Error number " & CStr(Err.Number) & " occurred in " & _
                Err.source & vbCrLf & vbCrLf & Err.Description

End Sub

'------------------------------------------------------------------------
'       MENU ITEMS EVENTS
'------------------------------------------------------------------------

Private Sub mnuForms_Click(index As Integer)
    Select Case index
        Case 0
            frmIntro.Show
        Case 1
            frmVacation.Show
    End Select
End Sub

Private Sub mnuFile_Click(index As Integer)
    Unload Me
End Sub


