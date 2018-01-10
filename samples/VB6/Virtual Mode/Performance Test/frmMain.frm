VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerListView 1.7 - Performance Test Sample"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill ListViews"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txtItems 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "1000"
      Top             =   6150
      Width           =   735
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   6120
      Width           =   2415
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwVirtual 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Virtual ListView"
      Top             =   0
      Width           =   5655
      _cx             =   9975
      _cy             =   10610
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   0   'False
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   2
      AutoSizeColumns =   0   'False
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   10
      BkImagePositionY=   10
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   1048575
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      EmptyMarkupTextAlignment=   1
      Enabled         =   -1  'True
      FilterChangedTimeout=   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   2
      GridLines       =   0   'False
      GroupFooterForeColor=   -2147483640
      GroupHeaderForeColor=   -2147483640
      GroupMarginBottom=   0
      GroupMarginLeft =   0
      GroupMarginRight=   0
      GroupMarginTop  =   12
      GroupSortOrder  =   0
      HeaderFullDragging=   -1  'True
      HeaderHotTracking=   0   'False
      HeaderHoverTime =   -1
      HeaderOLEDragImageStyle=   0
      HideLabels      =   0   'False
      HotForeColor    =   -1
      HotMousePointer =   0
      HotTracking     =   0   'False
      HotTrackingHoverTime=   -1
      HoverTime       =   -1
      IMEMode         =   -1
      IncludeHeaderInTabOrder=   0   'False
      InsertMarkColor =   0
      ItemActivationMode=   2
      ItemAlignment   =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      JustifyIconColumns=   0   'False
      LabelWrap       =   -1  'True
      MinItemRowsVisibleInGroups=   0
      MousePointer    =   0
      MultiSelect     =   -1  'True
      OLEDragImageStyle=   0
      OutlineColor    =   -2147483633
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      Regional        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   0   'False
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   0   'False
      ShowStateImages =   0   'False
      ShowSubItemImages=   0   'False
      SimpleSelect    =   0   'False
      SingleRow       =   0   'False
      SnapToGrid      =   0   'False
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TextBackColor   =   -1
      TileViewItemLines=   2
      TileViewLabelMarginBottom=   0
      TileViewLabelMarginLeft=   0
      TileViewLabelMarginRight=   0
      TileViewLabelMarginTop=   0
      TileViewSubItemForeColor=   -1
      TileViewTileHeight=   -1
      TileViewTileWidth=   -1
      ToolTips        =   3
      UnderlinedItems =   0
      UseMinColumnWidths=   -1  'True
      UseSystemFont   =   -1  'True
      UseWorkAreas    =   0   'False
      View            =   0
      VirtualMode     =   -1  'True
      EmptyMarkupText =   "frmMain.frx":0000
      FooterIntroText =   "frmMain.frx":0072
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwNormal 
      Height          =   6015
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   "Normal ListView"
      Top             =   0
      Width           =   6135
      _cx             =   10821
      _cy             =   10610
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   0   'False
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   2
      AutoSizeColumns =   0   'False
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   10
      BkImagePositionY=   10
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   1048575
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      EmptyMarkupTextAlignment=   1
      Enabled         =   -1  'True
      FilterChangedTimeout=   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   2
      GridLines       =   0   'False
      GroupFooterForeColor=   -2147483640
      GroupHeaderForeColor=   -2147483640
      GroupMarginBottom=   0
      GroupMarginLeft =   0
      GroupMarginRight=   0
      GroupMarginTop  =   12
      GroupSortOrder  =   0
      HeaderFullDragging=   -1  'True
      HeaderHotTracking=   0   'False
      HeaderHoverTime =   -1
      HeaderOLEDragImageStyle=   0
      HideLabels      =   0   'False
      HotForeColor    =   -1
      HotMousePointer =   0
      HotTracking     =   0   'False
      HotTrackingHoverTime=   -1
      HoverTime       =   -1
      IMEMode         =   -1
      IncludeHeaderInTabOrder=   0   'False
      InsertMarkColor =   0
      ItemActivationMode=   2
      ItemAlignment   =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      JustifyIconColumns=   0   'False
      LabelWrap       =   -1  'True
      MinItemRowsVisibleInGroups=   0
      MousePointer    =   0
      MultiSelect     =   -1  'True
      OLEDragImageStyle=   0
      OutlineColor    =   -2147483633
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      Regional        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   0   'False
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   0   'False
      ShowStateImages =   0   'False
      ShowSubItemImages=   0   'False
      SimpleSelect    =   0   'False
      SingleRow       =   0   'False
      SnapToGrid      =   0   'False
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TextBackColor   =   -1
      TileViewItemLines=   2
      TileViewLabelMarginBottom=   0
      TileViewLabelMarginLeft=   0
      TileViewLabelMarginRight=   0
      TileViewLabelMarginTop=   0
      TileViewSubItemForeColor=   -1
      TileViewTileHeight=   -1
      TileViewTileWidth=   -1
      ToolTips        =   3
      UnderlinedItems =   0
      UseMinColumnWidths=   -1  'True
      UseSystemFont   =   -1  'True
      UseWorkAreas    =   0   'False
      View            =   0
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmMain.frx":0092
      FooterIntroText =   "frmMain.frx":0104
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Fill time: 0 ms/0 ms"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   6240
      Width           =   1350
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ite&ms:"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   6240
      Width           =   465
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewIcons 
         Caption         =   "&Icons"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "&Small Icons"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuViewTiles 
         Caption         =   "&Tiles"
      End
      Begin VB.Menu mnuViewExtendedTiles 
         Caption         =   "E&xtended Tiles"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowGroups 
         Caption         =   "Show in &Groups"
      End
   End
   Begin VB.Menu mnuItemAlignment 
      Caption         =   "&ItemAlignment"
      Begin VB.Menu mnuItemAlignmentTop 
         Caption         =   "&Top"
      End
      Begin VB.Menu mnuItemAlignmentLeft 
         Caption         =   "&Left"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private bComctl32Version610OrNewer As Boolean
  Private bShowInGroups As Boolean
  Private cImages As Long
  Private cItems As Long
  Private hImgLst(1 To 3) As Long
  Private objActiveCtl As Object
  Private themeableOS As Boolean


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub cmdFill_Click()
  Dim eNormal As Long
  Dim eVirtual As Long
  Dim groupID As Long
  Dim iIcon As Long
  Dim iItem As Long
  Dim itm As ListViewItem
  Dim si1 As TILEVIEWSUBITEM
  Dim si2 As TILEVIEWSUBITEM
  Dim sNormal As Long
  Dim sVirtual As Long

  cmdFill.Enabled = False

  sVirtual = GetTickCount
  ExLVwVirtual.DontRedraw = True
  If bShowInGroups Then
    With ExLVwVirtual.Groups
      .RemoveAll
      .Add "Group 1", 1, , cItems \ 2, , , True
      .Add "Group 2", 2, , cItems - (cItems \ 2), , , True
    End With
  End If
  ExLVwVirtual.ShowGroups = bShowInGroups
  ExLVwVirtual.VirtualItemCount = cItems
  ExLVwVirtual.DontRedraw = False
  eVirtual = GetTickCount

  DoEvents
  DoEvents
  DoEvents

  sNormal = GetTickCount
  ExLVwNormal.DontRedraw = True
  si1.ColumnIndex = 1
  si1.WrapToMultipleLines = True
  si2.ColumnIndex = 2
  si2.BeginNewColumn = True
  si2.WrapToMultipleLines = True
  If bShowInGroups Then
    With ExLVwNormal.Groups
      .RemoveAll
      .Add "Group 1", 1, , , , , True
      .Add "Group 2", 2, , , , , True
    End With
  End If
  ExLVwNormal.ShowGroups = bShowInGroups
  With ExLVwNormal.ListItems
    .RemoveAll
    groupID = IIf(bShowInGroups, 1, -2)
    For iItem = 0 To cItems - 1
      If bShowInGroups And (iItem = cItems \ 2) Then groupID = 2
      Set itm = .Add("Item " & iItem, , iItem Mod cImages, , , groupID, Array(si1, si2))
      With itm.SubItems(1)
        .Text = "Item " & iItem & ", SubItem 1"
        .IconIndex = (iItem + 1) Mod cImages
      End With
      With itm.SubItems(2)
        .Text = "Item " & iItem & ", SubItem 2"
        .IconIndex = (iItem + 1) Mod cImages
      End With
    Next iItem
  End With
  ExLVwNormal.DontRedraw = False
  eNormal = GetTickCount

  lblTime.Caption = "Fill time: " & CStr(eVirtual - sVirtual) & " ms/" & CStr(eNormal - sNormal) & " ms"
  cmdFill.Enabled = True
End Sub

Private Sub ExLVwNormal_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumnsNormal
End Sub

Private Sub ExLVwNormal_GotFocus()
  Set objActiveCtl = ExLVwNormal
  UpdateMenu
End Sub

Private Sub ExLVwVirtual_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumnsVirtual
End Sub

Private Sub ExLVwVirtual_GotFocus()
  Set objActiveCtl = ExLVwVirtual
  UpdateMenu
End Sub

Private Sub ExLVwVirtual_ItemGetDisplayInfo(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal requestedInfo As ExLVwLibUCtl.RequestedInfoConstants, IconIndex As Long, Indent As Long, groupID As Long, TileViewColumns() As ExLVwLibUCtl.TILEVIEWSUBITEM, ByVal maxItemTextLength As Long, itemText As String, OverlayIndex As Long, StateImageIndex As Long, itemStates As ExLVwLibUCtl.ItemStateConstants, dontAskAgain As Boolean)
  Dim c(1 To 2) As TILEVIEWSUBITEM

  If requestedInfo And RequestedInfoConstants.riItemText Then
    If listSubItem Is Nothing Then
      itemText = "Item " & listItem.Index
    Else
      itemText = "Item " & listItem.Index & ", SubItem " & listSubItem.Index
    End If
  End If
  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = listItem.Index Mod cImages
  End If
  If requestedInfo And RequestedInfoConstants.riTileViewColumns Then
    c(1).ColumnIndex = 1
    c(1).WrapToMultipleLines = True
    c(2).ColumnIndex = 2
    c(2).BeginNewColumn = True
    c(2).WrapToMultipleLines = True
    TileViewColumns = c
  End If
End Sub

Private Sub ExLVwVirtual_MapGroupWideToTotalItemIndex(ByVal groupIndex As Long, ByVal groupWideItemIndex As Long, totalItemIndex As Long)
  totalItemIndex = groupIndex * (cItems \ 2) + groupWideItemIndex
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim hMod As Long
  Dim i As Long
  Dim iconsDir As String
  Dim iconPath As String
  Dim iconSize As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
    If bComctl32Version600OrNewer Then
      If .dwMajor = 6 Then
        bComctl32Version610OrNewer = (.dwMinor >= 10)
      Else
        bComctl32Version610OrNewer = True
      End If
    End If
  End With

  For i = 1 To 3
    iconSize = Choose(i, 16, 32, 48)
    hImgLst(i) = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
    If Right$(App.Path, 3) = "bin" Then
      iconsDir = App.Path & "\..\res\"
    Else
      iconsDir = App.Path & "\res\"
    End If
    iconsDir = iconsDir & iconSize & "x" & iconSize & IIf(bComctl32Version600OrNewer, "x32bpp\", "x8bpp\")
    iconPath = Dir$(iconsDir & "*.ico")
    While iconPath <> ""
      hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
      If hIcon Then
        ImageList_AddIcon hImgLst(i), hIcon
        DestroyIcon hIcon
      End If
      iconPath = Dir$
    Wend
  Next i
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  cItems = CLng(txtItems.Text)

  ExLVwVirtual.hImageList(ImageListConstants.ilSmall) = hImgLst(1)
  ExLVwVirtual.hImageList(ImageListConstants.ilLarge) = hImgLst(2)
  ExLVwVirtual.hImageList(ImageListConstants.ilExtraLarge) = hImgLst(3)
  ExLVwNormal.hImageList(ImageListConstants.ilSmall) = hImgLst(1)
  ExLVwNormal.hImageList(ImageListConstants.ilLarge) = hImgLst(2)
  ExLVwNormal.hImageList(ImageListConstants.ilExtraLarge) = hImgLst(3)
  cImages = ImageList_GetImageCount(hImgLst(1))

  Set objActiveCtl = ExLVwNormal
  mnuShowGroups.Enabled = bComctl32Version610OrNewer
  UpdateMenu
  InsertColumnsNormal
  InsertColumnsVirtual
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    cmdAbout.Move Me.ScaleWidth - cmdAbout.Width - 5, Me.ScaleHeight - cmdAbout.Height - 5
    lblDescr.Move 5, Me.ScaleHeight - lblDescr.Height - 10
    txtItems.Move lblDescr.Left + lblDescr.Width + 3, lblDescr.Top - 4
    cmdFill.Move txtItems.Left + txtItems.Width + 5, cmdAbout.Top
    lblTime.Move cmdFill.Left + cmdFill.Width + 10, lblDescr.Top
    ExLVwVirtual.Move 0, 0, Me.ScaleWidth / 2 - 3, cmdAbout.Top - 5
    ExLVwNormal.Move ExLVwVirtual.Left + ExLVwVirtual.Width + 6, 0, ExLVwVirtual.Width, ExLVwVirtual.Height
  End If
End Sub

Private Sub Form_Terminate()
  Dim i As Long

  For i = 1 To 3
    If hImgLst(i) Then ImageList_Destroy hImgLst(i)
  Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub mnuItemAlignmentLeft_Click()
  objActiveCtl.ItemAlignment = iaLeft
  UpdateMenu
End Sub

Private Sub mnuItemAlignmentTop_Click()
  objActiveCtl.ItemAlignment = iaTop
  UpdateMenu
End Sub

Private Sub mnuShowGroups_Click()
  mnuShowGroups.Checked = Not mnuShowGroups.Checked
  bShowInGroups = mnuShowGroups.Checked
End Sub

Private Sub mnuViewDetails_Click()
  objActiveCtl.View = vDetails
  objActiveCtl.FullRowSelect = frsExtendedMode
  UpdateMenu
End Sub

Private Sub mnuViewExtendedTiles_Click()
  objActiveCtl.View = vExtendedTiles
  objActiveCtl.FullRowSelect = frsDisabled
  UpdateMenu
End Sub

Private Sub mnuViewIcons_Click()
  objActiveCtl.View = vIcons
  UpdateMenu
End Sub

Private Sub mnuViewList_Click()
  objActiveCtl.View = vList
  UpdateMenu
End Sub

Private Sub mnuViewSmallIcons_Click()
  objActiveCtl.View = vSmallIcons
  UpdateMenu
End Sub

Private Sub mnuViewTiles_Click()
  objActiveCtl.View = vTiles
  objActiveCtl.FullRowSelect = frsDisabled
  UpdateMenu
End Sub

Private Sub txtItems_Change()
  On Error Resume Next
  cItems = CLng(txtItems.Text)
End Sub

Private Sub txtItems_LostFocus()
  txtItems.Text = cItems
End Sub


Private Sub InsertColumnsNormal()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLVwNormal.hWnd, StrPtr("explorer"), 0
  End If

  With ExLVwNormal.Columns
    .Add "Column 1", , 160, , , 1
    .Add "Column 2", , 160, , alCenter, 2
    .Add "Column 3", , 160, , alCenter, 3
  End With
End Sub

Private Sub InsertColumnsVirtual()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLVwVirtual.hWnd, StrPtr("explorer"), 0
  End If

  With ExLVwVirtual.Columns
    .Add "Column 1", , 160, , , 1
    .Add "Column 2", , 160, , alCenter, 2
    .Add "Column 3", , 160, , alCenter, 3
  End With
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong ExLVwVirtual.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ExLVwNormal.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Sub UpdateMenu()
  Select Case objActiveCtl.View
    Case vDetails
      mnuViewDetails.Checked = True
      mnuViewIcons.Checked = False
      mnuViewList.Checked = False
      mnuViewSmallIcons.Checked = False
      mnuViewTiles.Checked = False
      mnuViewExtendedTiles.Checked = False
    Case vIcons
      mnuViewDetails.Checked = False
      mnuViewIcons.Checked = True
      mnuViewList.Checked = False
      mnuViewSmallIcons.Checked = False
      mnuViewTiles.Checked = False
      mnuViewExtendedTiles.Checked = False
    Case vList
      mnuViewDetails.Checked = False
      mnuViewIcons.Checked = False
      mnuViewList.Checked = True
      mnuViewSmallIcons.Checked = False
      mnuViewTiles.Checked = False
      mnuViewExtendedTiles.Checked = False
    Case vSmallIcons
      mnuViewDetails.Checked = False
      mnuViewIcons.Checked = False
      mnuViewList.Checked = False
      mnuViewSmallIcons.Checked = True
      mnuViewTiles.Checked = False
      mnuViewExtendedTiles.Checked = False
    Case vTiles
      mnuViewDetails.Checked = False
      mnuViewIcons.Checked = False
      mnuViewList.Checked = False
      mnuViewSmallIcons.Checked = False
      mnuViewTiles.Checked = True
      mnuViewExtendedTiles.Checked = False
    Case vExtendedTiles
      mnuViewDetails.Checked = False
      mnuViewIcons.Checked = False
      mnuViewList.Checked = False
      mnuViewSmallIcons.Checked = False
      mnuViewTiles.Checked = False
      mnuViewExtendedTiles.Checked = True
  End Select
  mnuViewTiles.Enabled = bComctl32Version610OrNewer
  mnuViewExtendedTiles.Enabled = bComctl32Version610OrNewer
  Select Case objActiveCtl.ItemAlignment
    Case iaTop
      mnuItemAlignmentTop.Checked = True
      mnuItemAlignmentLeft.Checked = False
    Case iaLeft
      mnuItemAlignmentTop.Checked = False
      mnuItemAlignmentLeft.Checked = True
  End Select
End Sub
