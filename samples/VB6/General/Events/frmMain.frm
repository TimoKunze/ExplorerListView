VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Object = "{2B7096FB-1E68-4E54-A2D9-4AA933EAC680}#1.7#0"; "ExLVwA.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerListView 1.7 - Events Sample"
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
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   5400
      Width           =   975
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
      Left            =   8430
      TabIndex        =   4
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin ExLVwLibACtl.ExplorerListView ExLVwA 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   7215
      _cx             =   12726
      _cy             =   5106
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   2
      AutoSizeColumns =   0   'False
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   0
      BkImagePositionY=   0
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   0
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
      RegisterForOLEDragDrop=   -1  'True
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   -1  'True
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   -1  'True
      ShowStateImages =   -1  'True
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
      View            =   3
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmMain.frx":0000
      FooterIntroText =   "frmMain.frx":0072
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwU 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   4895
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
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
      ColumnHeaderVisibility=   2
      DisabledEvents  =   0
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
      RegisterForOLEDragDrop=   -1  'True
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   -1  'True
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   -1  'True
      ShowStateImages =   -1  'True
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
      View            =   3
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmMain.frx":00CC
      FooterIntroText =   "frmMain.frx":013E
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
      Begin VB.Menu mnuAutoArrange 
         Caption         =   "&Auto-arrange"
      End
      Begin VB.Menu mnuJustifyIconColumns 
         Caption         =   "&Justify Icon Columns"
      End
      Begin VB.Menu mnuSnapToGrid 
         Caption         =   "Snap to G&rid"
      End
      Begin VB.Menu mnuShowGroups 
         Caption         =   "Show in &Groups"
      End
      Begin VB.Menu mnuCheckOnSelect 
         Caption         =   "&Check on Select"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoSizeColumns 
         Caption         =   "Auto-Si&ze Columns"
      End
      Begin VB.Menu mnuResizableColumns 
         Caption         =   "R&esizable Columns"
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


  Private Const MAX_PATH = 260


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type POINTAPI
    x As Long
    y As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private bComctl32Version610OrNewer As Boolean
  Private bLog As Boolean
  Private draggedPictureA As StdPicture
  Private draggedPictureU As StdPicture
  Private hBMPDownArrow As Long
  Private hBMPUpArrow As Long
  Private hImgLst(1 To 4) As Long
  Private objActiveCtl As Object
  Private ptStartDrag As POINTAPI
  Private themeableOS As Boolean
  Private wallpaperA As StdPicture
  Private wallpaperU As StdPicture

  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function PropVariantToString Lib "propsys.dll" (ByVal propvar As Long, ByVal psz As Long, ByVal cch As Long) As Long
  Private Declare Function PSStringFromPropertyKey Lib "propsys.dll" (ByVal pkey As Long, ByVal psz As Long, ByVal cch As Long) As Long
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


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub ExLVwA_AbortedDrag()
  AddLogEntry "ExLVwA_AbortedDrag"
  Set ExLVwA.DropHilitedItem = Nothing
End Sub

Private Sub ExLVwA_AfterScroll(ByVal dx As Long, ByVal dy As Long)
  AddLogEntry "ExLVwA_AfterScroll: dx=" & dx & ", dy=" & dy
End Sub

Private Sub ExLVwA_BeforeScroll(ByVal dx As Long, ByVal dy As Long)
  AddLogEntry "ExLVwA_BeforeScroll: dx=" & dx & ", dy=" & dy
End Sub

Private Sub ExLVwA_BeginColumnResizing(ByVal column As ExLVwLibACtl.IListViewColumn, Cancel As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_BeginColumnResizing: column=Nothing, cancel=" & Cancel
  Else
    AddLogEntry "ExLVwA_BeginColumnResizing: column=" & column.Caption & ", cancel=" & Cancel
  End If
End Sub

Private Sub ExLVwA_BeginMarqueeSelection(Cancel As Boolean)
  AddLogEntry "ExLVwA_BeginMarqueeSelection: cancel=" & Cancel
End Sub

Private Sub ExLVwA_CacheItemsHint(ByVal FirstItem As ExLVwLibACtl.IListViewItem, ByVal lastItem As ExLVwLibACtl.IListViewItem)
  AddLogEntry "ExLVwA_CacheItemsHint: firstItem=" & FirstItem.Index & ", lastItem=" & lastItem.Index
End Sub

Private Sub ExLVwA_CancelSubItemEdit(ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal editMode As ExLVwLibACtl.SubItemEditModeConstants)
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwA_CancelSubItemEdit: listItem=Nothing, listSubItem=Nothing, editMode=" & editMode
  Else
    AddLogEntry "ExLVwA_CancelSubItemEdit: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", editMode=" & editMode
  End If
End Sub

Private Sub ExLVwA_CaretChanged(ByVal previousCaretItem As ExLVwLibACtl.IListViewItem, ByVal newCaretItem As ExLVwLibACtl.IListViewItem)
  Dim str As String

  str = "ExLVwA_CaretChanged: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwA_ChangedSortOrder(ByVal previousSortOrder As ExLVwLibACtl.SortOrderConstants, ByVal newSortOrder As ExLVwLibACtl.SortOrderConstants)
  AddLogEntry "ExLVwA_ChangedSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder
End Sub

Private Sub ExLVwA_ChangedWorkAreas(ByVal WorkAreas As ExLVwLibACtl.IListViewWorkAreas)
  If WorkAreas Is Nothing Then
    AddLogEntry "ExLVwA_ChangedWorkAreas: WorkAreas=Nothing"
  Else
    AddLogEntry "ExLVwA_ChangedWorkAreas: WorkAreas=" & WorkAreas.Count
  End If
End Sub

Private Sub ExLVwA_ChangingSortOrder(ByVal previousSortOrder As ExLVwLibACtl.SortOrderConstants, ByVal newSortOrder As ExLVwLibACtl.SortOrderConstants, cancelChange As Boolean)
  AddLogEntry "ExLVwA_ChangingSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder & ", cancelChange=" & cancelChange
End Sub

Private Sub ExLVwA_ChangingWorkAreas(ByVal WorkAreas As ExLVwLibACtl.IVirtualListViewWorkAreas, cancelChanges As Boolean)
  If WorkAreas Is Nothing Then
    AddLogEntry "ExLVwA_ChangingWorkAreas: WorkAreas=Nothing, cancelChanges=" & cancelChanges
  Else
    AddLogEntry "ExLVwA_ChangingWorkAreas: WorkAreas=" & WorkAreas.Count & ", cancelChanges=" & cancelChanges
  End If
End Sub

Private Sub ExLVwA_Click(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_Click: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_CollapsedGroup(ByVal Group As ExLVwLibACtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_CollapsedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwA_CollapsedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_ColumnBeginDrag(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, doAutomaticDragDrop As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnBeginDrag: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDragDrop=" & doAutomaticDragDrop
  Else
    AddLogEntry "ExLVwA_ColumnBeginDrag: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDragDrop=" & doAutomaticDragDrop
  End If
End Sub

Private Sub ExLVwA_ColumnClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  Dim col As ExLVwLibACtl.ListViewColumn
  Dim curArrow As ExLVwLibACtl.SortArrowConstants
  Dim newArrow As ExLVwLibACtl.SortArrowConstants

  If column Is Nothing Then
    ' shouldn't happen
    AddLogEntry "ExLVwA_ColumnClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_ColumnClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

    ' update the sort arrows
    For Each col In ExLVwA.Columns
      If col.Index = column.Index Then
        If bComctl32Version600OrNewer Then
          curArrow = col.SortArrow
        Else
          Select Case col.BitmapHandle
            Case 0
              curArrow = saNone
            Case hBMPDownArrow
              curArrow = saDown
            Case hBMPUpArrow
              curArrow = saUp
          End Select
        End If

        If curArrow = saUp Then
          newArrow = saDown
        Else
          newArrow = saUp
        End If
      Else
        newArrow = saNone
      End If

      If bComctl32Version600OrNewer Then
        col.SortArrow = newArrow
        Select Case newArrow
          Case saDown
            ExLVwA.SortOrder = soDescending
            Set ExLVwA.SelectedColumn = col
            ExLVwA.SortItems , , , , , col, True
          Case saUp
            ExLVwA.SortOrder = soAscending
            Set ExLVwA.SelectedColumn = col
            ExLVwA.SortItems , , , , , col, True
        End Select
      Else
        Select Case newArrow
          Case saNone
            col.BitmapHandle = 0
          Case saDown
            col.BitmapHandle = hBMPDownArrow
            ExLVwA.SortOrder = soDescending
            ExLVwA.SortItems , , , , , col
          Case saUp
            col.BitmapHandle = hBMPUpArrow
            ExLVwA.SortOrder = soAscending
            ExLVwA.SortItems , , , , , col
        End Select
      End If
    Next col
  End If
End Sub

Private Sub ExLVwA_ColumnDropDown(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnDropDown: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExLVwA_ColumnDropDown: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu

    ' TODO: Display drop-down
  End If
End Sub

Private Sub ExLVwA_ColumnEndAutoDragDrop(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, doAutomaticDrop As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnEndAutoDragDrop: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDrop=" & doAutomaticDrop
  Else
    AddLogEntry "ExLVwA_ColumnEndAutoDragDrop: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDrop=" & doAutomaticDrop
  End If
End Sub

Private Sub ExLVwA_ColumnMouseEnter(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnMouseEnter: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_ColumnMouseEnter: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_ColumnMouseLeave(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnMouseLeave: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_ColumnMouseLeave: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_ColumnStateImageChanged(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExLVwLibACtl.StateImageChangeCausedByConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnStateImageChanged: column=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  Else
    AddLogEntry "ExLVwA_ColumnStateImageChanged: column=" & column.Caption & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy

    If column.Index = 0 Then
      ExLVwA.ListItems(-1).StateImageIndex = column.StateImageIndex
    End If
  End If
End Sub

Private Sub ExLVwA_ColumnStateImageChanging(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExLVwLibACtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ColumnStateImageChanging: column=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  Else
    AddLogEntry "ExLVwA_ColumnStateImageChanging: column=" & column.Caption & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExLVwA_CompareGroups(ByVal firstGroup As ExLVwLibACtl.IListViewGroup, ByVal secondGroup As ExLVwLibACtl.IListViewGroup, result As ExLVwLibACtl.CompareResultConstants)
  Dim str As String

  str = "ExLVwA_CompareGroups: firstGroup="
  If firstGroup Is Nothing Then
    str = str & "Nothing, secondGroup="
  Else
    str = str & firstGroup.Text & ", secondGroup="
  End If
  If secondGroup Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondGroup.Text
  End If
  str = str & ", result=" & result

  AddLogEntry str
End Sub

Private Sub ExLVwA_CompareItems(ByVal firstItem As ExLVwLibACtl.IListViewItem, ByVal secondItem As ExLVwLibACtl.IListViewItem, result As ExLVwLibACtl.CompareResultConstants)
  Dim str As String

  str = "ExLVwA_CompareItems: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & firstItem.Text & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Text
  End If
  str = str & ", result=" & result

  AddLogEntry str
End Sub

Private Sub ExLVwA_ConfigureSubItemControl(ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal controlKind As ExLVwLibACtl.SubItemControlKindConstants, ByVal editMode As ExLVwLibACtl.SubItemEditModeConstants, ByVal subItemControl As ExLVwLibACtl.SubItemControlConstants, themeAppName As String, themeIDList As String, hFont As Long, textColor As stdole.OLE_COLOR, pPropertyDescription As Long, ByVal pPropertyValue As Long)
  Dim buffer As String

  buffer = String$(1024, 0)
  PropVariantToString pPropertyValue, StrPtr(buffer), 1024
  buffer = Left$(buffer, lstrlen(StrPtr(buffer)))
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwA_ConfigureSubItemControl: listItem=Nothing, listSubItem=Nothing, controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl & ", themeAppName=" & themeAppName & ", themeIDList=" & themeIDList & ", hFont=0x" & Hex(hFont) & ", textColor=0x" & Hex(textColor) & ", pPropertyDescription=0x" & Hex(pPropertyDescription) & ", propertyValue=" & buffer
  Else
    AddLogEntry "ExLVwA_ConfigureSubItemControl: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl & ", themeAppName=" & themeAppName & ", themeIDList=" & themeIDList & ", hFont=0x" & Hex(hFont) & ", textColor=0x" & Hex(textColor) & ", pPropertyDescription=0x" & Hex(pPropertyDescription) & ", propertyValue=" & buffer
  End If
End Sub

Private Sub ExLVwA_ContextMenu(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  Dim str As String

  str = "ExLVwA_ContextMenu: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu

  AddLogEntry str

  If listItem Is Nothing Then
    PopupMenu mnuView, , ExLVwA.Left + ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ExLVwA.Top + ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  Else
    '
  End If
End Sub

Private Sub ExLVwA_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExLVwA_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExLVwA_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  AddLogEntry "ExLVwA_CreatedHeaderControlWindow: hWndHeader=0x" & Hex(hWndHeader)
  InsertColumnsA
End Sub

Private Sub ExLVwA_CustomDraw(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal drawAllItems As Boolean, textColor As stdole.OLE_COLOR, TextBackColor As stdole.OLE_COLOR, ByVal drawStage As ExLVwLibACtl.CustomDrawStageConstants, ByVal itemState As ExLVwLibACtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibACtl.RECTANGLE, furtherProcessing As ExLVwLibACtl.CustomDrawReturnValuesConstants)
'  Dim str As String
'
'  str = "ExLVwA_CustomDraw: listItem="
'  If listItem Is Nothing Then
'    str = str & "Nothing, listSubItem="
'  Else
'    str = str & listItem.Text & ", listSubItem="
'  End If
'  If listSubItem Is Nothing Then
'    str = str & "Nothing"
'  Else
'    str = str & listSubItem.Text
'  End If
'  str = str & ", drawAllItems=" & drawAllItems & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(TextBackColor) & ", drawStage=" & drawStage & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'
'  AddLogEntry str
End Sub

Private Sub ExLVwA_DblClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_DblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExLVwA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ExLVwA_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExLVwA_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExLVwA_DestroyedHeaderControlWindow(ByVal hWndHeader As Long)
  AddLogEntry "ExLVwA_DestroyedHeaderControlWindow: hWndHeader=0x" & Hex(hWndHeader)
End Sub

Private Sub ExLVwA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ExLVwA_DragDrop"
End Sub

Private Sub ExLVwA_DragMouseMove(dropTarget As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwA_DragMouseMove: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  Else
    AddLogEntry "ExLVwA_DragMouseMove: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  End If
End Sub

Private Sub ExLVwA_DragOver(Source As Control, x As Single, y As Single, state As Integer)
  AddLogEntry "ExLVwA_DragOver"
End Sub

Private Sub ExLVwA_Drop(ByVal dropTarget As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim dx As Long
  Dim dy As Long
  Dim itm As ExLVwLibACtl.ListViewItem
  Dim xPos As Long
  Dim yPos As Long

  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwA_Drop: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_Drop: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If

  x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  dx = x - ptStartDrag.x
  dy = y - ptStartDrag.y

  For Each itm In ExLVwA.DraggedItems
    itm.GetPosition xPos, yPos
    itm.SetPosition xPos + dx, yPos + dy
  Next itm
End Sub

Private Sub ExLVwA_EditChange()
  AddLogEntry "ExLVwA_EditChange - " & ExLVwA.EditText
End Sub

Private Sub ExLVwA_EditClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "ExLVwA_EditContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub ExLVwA_EditDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditGotFocus()
  AddLogEntry "ExLVwA_EditGotFocus"
End Sub

Private Sub ExLVwA_EditKeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwA_EditKeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwA_EditKeyPress(keyAscii As Integer)
  AddLogEntry "ExLVwA_EditKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExLVwA_EditKeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwA_EditKeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwA_EditLostFocus()
  AddLogEntry "ExLVwA_EditLostFocus"
End Sub

Private Sub ExLVwA_EditMClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
'  AddLogEntry "ExLVwA_EditMouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditMouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditMouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As ExLVwLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "ExLVwA_EditMouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub ExLVwA_EditRClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditRClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditRDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditRDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditXClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditXClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EditXDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwA_EditXDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwA_EmptyMarkupTextLinkClick(ByVal linkIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  AddLogEntry "ExLVwA_EmptyMarkupTextLinkClick: linkIndex=" & linkIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ExLVwA_EndColumnResizing(ByVal column As ExLVwLibACtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_EndColumnResizing: column=Nothing"
  Else
    AddLogEntry "ExLVwA_EndColumnResizing: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwA_EndSubItemEdit(ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal editMode As ExLVwLibACtl.SubItemEditModeConstants, ByVal pPropertyKey As Long, ByVal pPropertyValue As Long, Cancel As Boolean)
  Dim propKey As String
  Dim propValue As String

  propKey = String$(1024, 0)
  PSStringFromPropertyKey pPropertyKey, StrPtr(propKey), 1024
  propKey = Left$(propKey, lstrlen(StrPtr(propKey)))
  propValue = String$(1024, 0)
  PropVariantToString pPropertyValue, StrPtr(propValue), 1024
  propValue = Left$(propValue, lstrlen(StrPtr(propValue)))
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwA_EndSubItemEdit: listItem=Nothing, listSubItem=Nothing, editMode=" & editMode & ", pPropertyKey=" & propKey & ", propertyValue=" & propValue & ", Cancel=" & Cancel
  Else
    AddLogEntry "ExLVwA_EndSubItemEdit: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", editMode=" & editMode & ", pPropertyKey=" & propKey & ", propertyValue=" & propValue & ", Cancel=" & Cancel
  End If
End Sub

Private Sub ExLVwA_ExpandedGroup(ByVal Group As ExLVwLibACtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_ExpandedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwA_ExpandedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_FilterButtonClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, filterButtonRectangle As ExLVwLibACtl.RECTANGLE, raiseFilterChanged As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_FilterButtonClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", filterButtonRectangle=(" & filterButtonRectangle.Left & "," & filterButtonRectangle.Top & ")-(" & filterButtonRectangle.Right & "," & filterButtonRectangle.Bottom & "), raiseFilterChanged=" & raiseFilterChanged
  Else
    AddLogEntry "ExLVwA_FilterButtonClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", filterButtonRectangle=(" & filterButtonRectangle.Left & "," & filterButtonRectangle.Top & ")-(" & filterButtonRectangle.Right & "," & filterButtonRectangle.Bottom & "), raiseFilterChanged=" & raiseFilterChanged
  End If
End Sub

Private Sub ExLVwA_FilterChanged(ByVal column As ExLVwLibACtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_FilterChanged: column=Nothing"
  Else
    AddLogEntry "ExLVwA_FilterChanged: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwA_FindVirtualItem(ByVal itemToStartWith As ExLVwLibACtl.IListViewItem, ByVal searchMode As ExLVwLibACtl.SearchModeConstants, searchFor As Variant, ByVal searchDirection As ExLVwLibACtl.SearchDirectionConstants, ByVal wrapAtLastItem As Boolean, foundItem As ExLVwLibACtl.IListViewItem)
  Dim str As String

  str = "ExLVwA_FindVirtualItem: itemToStartWith="
  If itemToStartWith Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & itemToStartWith.Text
  End If
  str = str & ", searchMode=" & searchMode & ", VarType(searchFor)=" & VarType(searchFor) & ", searchDirection=" & searchDirection & ", wrapAtLastItem=" & wrapAtLastItem & ", foundItem="
  If foundItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & foundItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwA_FooterItemClick(ByVal footerItem As ExLVwLibACtl.IListViewFooterItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants, removeFooterArea As Boolean)
  If footerItem Is Nothing Then
    AddLogEntry "ExLVwA_FooterItemClick: footerItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", removeFooterArea=" & removeFooterArea
  Else
    AddLogEntry "ExLVwA_FooterItemClick: footerItem=" & footerItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", removeFooterArea=" & removeFooterArea

    Select Case footerItem.ItemData
      Case 1
        MsgBox "Good night!"
      Case 2
        MsgBox "Let's go, Pinky!"
    End Select
  End If
End Sub

Private Sub ExLVwA_FreeColumnData(ByVal column As ExLVwLibACtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_FreeColumnData: column=Nothing"
  Else
    AddLogEntry "ExLVwA_FreeColumnData: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwA_FreeFooterItemData(ByVal footerItem As ExLVwLibACtl.IListViewFooterItem, ByVal ItemData As Long)
  If footerItem Is Nothing Then
    AddLogEntry "ExLVwA_FreeFooterItemData: footerItem=Nothing, itemData=" & ItemData
  Else
    AddLogEntry "ExLVwA_FreeFooterItemData: footerItem=" & footerItem.Text & ", itemData=" & ItemData
  End If
End Sub

Private Sub ExLVwA_FreeItemData(ByVal listItem As ExLVwLibACtl.IListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_FreeItemData: listItem=Nothing"
  Else
    AddLogEntry "ExLVwA_FreeItemData: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwA_GetSubItemControl(ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal controlKind As ExLVwLibACtl.SubItemControlKindConstants, ByVal editMode As ExLVwLibACtl.SubItemEditModeConstants, subItemControl As ExLVwLibACtl.SubItemControlConstants)
  If listSubItem Is Nothing Then
'    AddLogEntry "ExLVwA_GetSubItemControl: listItem=Nothing, listSubItem=Nothing, controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl
  Else
'    AddLogEntry "ExLVwA_GetSubItemControl: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl
  End If
End Sub

Private Sub ExLVwA_GotFocus()
  AddLogEntry "ExLVwA_GotFocus"
  Set objActiveCtl = ExLVwA
  UpdateMenu
End Sub

Private Sub ExLVwA_GroupAsynchronousDrawFailed(ByVal Group As ExLVwLibACtl.IListViewGroup, imageDetails As ExLVwLibACtl.FAILEDIMAGEDETAILS, ByVal failureReason As ExLVwLibACtl.ImageDrawingFailureReasonConstants, furtherProcessing As ExLVwLibACtl.FailedAsyncDrawReturnValuesConstants, newImageToDraw As Long)
  Dim str As String

  str = "ExLVwA_GroupAsynchronousDrawFailed: group="
  If Group Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Group.Text
  End If
  str = str & ", failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  AddLogEntry "     imageDetails.hImageList=0x" & Hex(imageDetails.hImageList)
  AddLogEntry "     imageDetails.hDC=0x" & Hex(imageDetails.hDC)
  AddLogEntry "     imageDetails.IconIndex=" & imageDetails.IconIndex
  AddLogEntry "     imageDetails.OverlayIndex=" & imageDetails.OverlayIndex
  AddLogEntry "     imageDetails.DrawingStyle=" & imageDetails.DrawingStyle
  AddLogEntry "     imageDetails.DrawingEffects=" & imageDetails.DrawingEffects
  AddLogEntry "     imageDetails.BackColor=0x" & Hex(imageDetails.BackColor)
  AddLogEntry "     imageDetails.ForeColor=0x" & Hex(imageDetails.ForeColor)
  AddLogEntry "     imageDetails.EffectColor=0x" & Hex(imageDetails.EffectColor)
End Sub

Private Sub ExLVwA_GroupCustomDraw(ByVal Group As ExLVwLibACtl.IListViewGroup, textColor As stdole.OLE_COLOR, headerAlignment As ExLVwLibACtl.AlignmentConstants, footerAlignment As ExLVwLibACtl.AlignmentConstants, ByVal drawStage As ExLVwLibACtl.CustomDrawStageConstants, ByVal groupState As ExLVwLibACtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibACtl.RECTANGLE, textRectangle As ExLVwLibACtl.RECTANGLE, furtherProcessing As ExLVwLibACtl.CustomDrawReturnValuesConstants)
'  If Group Is Nothing Then
'    AddLogEntry "ExLVwA_GroupCustomDraw: Group=Nothing, textColor=0x" & Hex(textColor) & ", headerAlignment=" & headerAlignment & ", footerAlignment=" & footerAlignment & ", drawStage=" & drawStage & ", groupState=" & groupState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ExLVwA_GroupCustomDraw: Group=" & Group.Text & ", textColor=0x" & Hex(textColor) & ", headerAlignment=" & headerAlignment & ", footerAlignment=" & footerAlignment & ", drawStage=" & drawStage & ", groupState=" & groupState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ExLVwA_GroupGotFocus(ByVal Group As ExLVwLibACtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_GroupGotFocus: group=Nothing"
  Else
    AddLogEntry "ExLVwA_GroupGotFocus: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_GroupLostFocus(ByVal Group As ExLVwLibACtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_GroupLostFocus: group=Nothing"
  Else
    AddLogEntry "ExLVwA_GroupLostFocus: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_GroupSelectionChanged(ByVal Group As ExLVwLibACtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_GroupSelectionChanged: group=Nothing"
  Else
    AddLogEntry "ExLVwA_GroupSelectionChanged: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_GroupTaskLinkClick(ByVal Group As ExLVwLibACtl.IListViewGroup, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_GroupTaskLinkClick: group=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_GroupTaskLinkClick: group=" & Group.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderAbortedDrag()
  AddLogEntry "ExLVwA_HeaderAbortedDrag"
End Sub

Private Sub ExLVwA_HeaderChevronClick(ByVal firstOverflownColumn As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  If firstOverflownColumn Is Nothing Then
    AddLogEntry "ExLVwA_HeaderChevronClick: firstOverflownColumn=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExLVwA_HeaderChevronClick: firstOverflownColumn=" & firstOverflownColumn.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
  End If
End Sub

Private Sub ExLVwA_HeaderClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderContextMenu(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, showDefaultMenu As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderContextMenu: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExLVwA_HeaderContextMenu: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
  End If
End Sub

Private Sub ExLVwA_HeaderCustomDraw(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal drawStage As ExLVwLibACtl.CustomDrawStageConstants, ByVal columnState As ExLVwLibACtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibACtl.RECTANGLE, furtherProcessing As ExLVwLibACtl.CustomDrawReturnValuesConstants)
'  If column Is Nothing Then
'    AddLogEntry "ExLVwA_HeaderCustomDraw: column=Nothing, drawStage=" & drawStage & ", columnState=" & columnState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ExLVwA_HeaderCustomDraw: column=" & column.Caption & ", drawStage=" & drawStage & ", columnState=" & columnState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ExLVwA_HeaderDblClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderDividerDblClick(ByVal column As ExLVwLibACtl.IListViewColumn, autoSizeColumn As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderDividerDblClick: column=Nothing, autoSizeColumn=" & autoSizeColumn
  Else
    AddLogEntry "ExLVwA_HeaderDividerDblClick: column=" & column.Caption & ", autoSizeColumn=" & autoSizeColumn
  End If
End Sub

Private Sub ExLVwA_HeaderDragMouseMove(dropTarget As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwA_HeaderDragMouseMove: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  Else
    AddLogEntry "ExLVwA_HeaderDragMouseMove: dropTarget=" & dropTarget.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  End If
End Sub

Private Sub ExLVwA_HeaderDrop(ByVal dropTarget As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwA_HeaderDrop: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderDrop: dropTarget=" & dropTarget.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderGotFocus()
  AddLogEntry "ExLVwA_HeaderGotFocus"
End Sub

Private Sub ExLVwA_HeaderItemGetDisplayInfo(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal requestedInfo As ExLVwLibACtl.RequestedInfoConstants, IconIndex As Long, ByVal maxColumnCaptionLength As Long, columnCaption As String, dontAskAgain As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderItemGetDisplayInfo: column=Nothing, requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", maxColumnCaptionLength=" & maxColumnCaptionLength & ", columnCaption=" & columnCaption & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "ExLVwA_HeaderItemGetDisplayInfo: column=" & column.Index & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", maxColumnCaptionLength=" & maxColumnCaptionLength & ", columnCaption=" & columnCaption & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub ExLVwA_HeaderKeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwA_HeaderKeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwA_HeaderKeyPress(keyAscii As Integer)
  AddLogEntry "ExLVwA_HeaderKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExLVwA_HeaderKeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwA_HeaderKeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwA_HeaderLostFocus()
  AddLogEntry "ExLVwA_HeaderLostFocus"
End Sub

Private Sub ExLVwA_HeaderMClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMDblClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseDown(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMouseDown: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMouseDown: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseEnter(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMouseEnter: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMouseEnter: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseHover(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMouseHover: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMouseHover: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseLeave(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMouseLeave: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMouseLeave: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseMove(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
'    AddLogEntry "ExLVwA_HeaderMouseMove: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
'    AddLogEntry "ExLVwA_HeaderMouseMove: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseUp(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMouseUp: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderMouseUp: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderMouseWheel(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, ByVal scrollAxis As ExLVwLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderMouseWheel: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
  Else
    AddLogEntry "ExLVwA_HeaderMouseWheel: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
  End If
End Sub

Private Sub ExLVwA_HeaderOLECompleteDrag(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal performedEffect As ExLVwLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwA_HeaderOLEDragDrop(ByVal data As ExLVwLibACtl.IOLEDataObject, effect As ExLVwLibACtl.OLEDropEffectConstants, ByVal dropTarget As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFDIB) Then
    Set draggedPictureA = data.GetData(vbCFDIB)
    ExLVwA.BkImage = draggedPictureA.Handle
  End If
  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ExLVwA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwA_HeaderOLEDragEnter(ByVal data As ExLVwLibACtl.IOLEDataObject, effect As ExLVwLibACtl.OLEDropEffectConstants, dropTarget As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderOLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ExLVwA_HeaderOLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ExLVwA_HeaderOLEDragLeave(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal dropTarget As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderOLEDragLeavePotentialTarget()
  AddLogEntry "ExLVwA_HeaderOLEDragLeavePotentialTarget"
End Sub

Private Sub ExLVwA_HeaderOLEDragMouseMove(ByVal data As ExLVwLibACtl.IOLEDataObject, effect As ExLVwLibACtl.OLEDropEffectConstants, dropTarget As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderOLEGiveFeedback(ByVal effect As ExLVwLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ExLVwA_HeaderOLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ExLVwA_HeaderOLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As ExLVwLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "ExLVwA_HeaderOLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ExLVwA_HeaderOLEReceivedNewData(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderOLESetData(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderOLEStartDrag(ByVal data As ExLVwLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_HeaderOLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderOwnerDrawItem(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal itemState As ExLVwLibACtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibACtl.RECTANGLE)
  Dim str As String

  str = "ExLVwA_HeaderOwnerDrawItem: column="
  If column Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & column.Caption
  End If
  str = str & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub ExLVwA_HeaderRClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderRClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderRClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderRDblClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderRDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderRDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderXClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderXClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderXClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HeaderXDblClick(ByVal column As ExLVwLibACtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_HeaderXDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_HeaderXDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_HotItemChanged(ByVal previousHotItem As ExLVwLibACtl.IListViewItem, ByVal newHotItem As ExLVwLibACtl.IListViewItem)
  Dim str As String

  str = "ExLVwA_HotItemChanged: previousHotItem="
  If previousHotItem Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotItem.Text & ", newHotItem="
  End If
  If newHotItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwA_HotItemChanging(ByVal previousHotItem As ExLVwLibACtl.IListViewItem, ByVal newHotItem As ExLVwLibACtl.IListViewItem, cancelChange As Boolean)
  Dim str As String

  str = "ExLVwA_HotItemChanging: previousHotItem="
  If previousHotItem Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotItem.Text & ", newHotItem="
  End If
  If newHotItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotItem.Text
  End If
  str = str & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ExLVwA_IncrementalSearching(ByVal currentSearchString As String, itemToSelect As Long)
  AddLogEntry "ExLVwA_IncrementalSearching: currentSearchString=" & currentSearchString & ", itemToSelect=" & itemToSelect
End Sub

Private Sub ExLVwA_IncrementalSearchStringChanging(ByVal currentSearchString As String, ByVal keyCodeOfCharToBeAdded As Integer, cancelChange As Boolean)
  AddLogEntry "ExLVwA_IncrementalSearchStringChanging: currentSearchString=" & currentSearchString & ", keyCodeOfCharToBeAdded=" & keyCodeOfCharToBeAdded & ", cancelChange=" & cancelChange
End Sub

Private Sub ExLVwA_InsertedColumn(ByVal column As ExLVwLibACtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_InsertedColumn: column=Nothing"
  Else
    AddLogEntry "ExLVwA_InsertedColumn: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwA_InsertedGroup(ByVal Group As ExLVwLibACtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_InsertedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwA_InsertedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_InsertedItem(ByVal listItem As ExLVwLibACtl.IListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_InsertedItem: listItem=Nothing"
  Else
    AddLogEntry "ExLVwA_InsertedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwA_InsertingColumn(ByVal column As ExLVwLibACtl.IVirtualListViewColumn, cancelInsertion As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_InsertingColumn: column=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExLVwA_InsertingColumn: column=" & column.Caption & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExLVwA_InsertingGroup(ByVal Group As ExLVwLibACtl.IVirtualListViewGroup, cancelInsertion As Boolean)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_InsertingGroup: group=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExLVwA_InsertingGroup: group=" & Group.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExLVwA_InsertingItem(ByVal listItem As ExLVwLibACtl.IVirtualListViewItem, cancelInsertion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_InsertingItem: listItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExLVwA_InsertingItem: listItem=" & listItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExLVwA_InvokeVerbFromSubItemControl(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal verb As String)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_InvokeVerbFromSubItemControl: listItem=Nothing, verb=" & verb
  Else
    AddLogEntry "ExLVwA_InvokeVerbFromSubItemControl: listItem=" & listItem.Text & ", verb=" & verb
  End If
End Sub

Private Sub ExLVwA_ItemActivate(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal shift As Integer, ByVal x As Long, ByVal y As Long)
  Dim str As String

  str = "ExLVwA_ItemActivate: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ExLVwA_ItemAsynchronousDrawFailed(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, imageDetails As ExLVwLibACtl.FAILEDIMAGEDETAILS, ByVal failureReason As ExLVwLibACtl.ImageDrawingFailureReasonConstants, furtherProcessing As ExLVwLibACtl.FailedAsyncDrawReturnValuesConstants, newImageToDraw As Long)
  Dim str As String

  str = "ExLVwA_ItemAsynchronousDrawFailed: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  AddLogEntry "     imageDetails.hImageList=0x" & Hex(imageDetails.hImageList)
  AddLogEntry "     imageDetails.hDC=0x" & Hex(imageDetails.hDC)
  AddLogEntry "     imageDetails.IconIndex=" & imageDetails.IconIndex
  AddLogEntry "     imageDetails.OverlayIndex=" & imageDetails.OverlayIndex
  AddLogEntry "     imageDetails.DrawingStyle=" & imageDetails.DrawingStyle
  AddLogEntry "     imageDetails.DrawingEffects=" & imageDetails.DrawingEffects
  AddLogEntry "     imageDetails.BackColor=0x" & Hex(imageDetails.BackColor)
  AddLogEntry "     imageDetails.ForeColor=0x" & Hex(imageDetails.ForeColor)
  AddLogEntry "     imageDetails.EffectColor=0x" & Hex(imageDetails.EffectColor)
End Sub

Private Sub ExLVwA_ItemBeginDrag(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim col As ExLVwLibACtl.ListViewItems
  Dim str As String

  str = "ExLVwA_ItemBeginDrag: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  On Error Resume Next
  Set col = ExLVwA.ListItems
  col.FilterType(ExLVwLibACtl.FilteredPropertyConstants.fpSelected) = ExLVwLibACtl.FilterTypeConstants.ftIncluding
  col.Filter(ExLVwLibACtl.FilteredPropertyConstants.fpSelected) = Array(True)

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ExLVwA.BeginDrag ExLVwA.CreateItemContainer(col), -1
End Sub

Private Sub ExLVwA_ItemBeginRDrag(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_ItemBeginRDrag: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_ItemGetDisplayInfo(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal requestedInfo As ExLVwLibACtl.RequestedInfoConstants, IconIndex As Long, Indent As Long, groupID As Long, TileViewColumns() As ExLVwLibACtl.TILEVIEWSUBITEM, ByVal maxItemTextLength As Long, itemText As String, OverlayIndex As Long, StateImageIndex As Long, itemStates As ExLVwLibACtl.ItemStateConstants, dontAskAgain As Boolean)
  Dim c(1 To 2) As ExLVwLibACtl.TILEVIEWSUBITEM
  Dim str As String

  str = "ExLVwA_ItemGetDisplayInfo: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Index & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Index
  End If
  str = str & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", Indent=" & Indent & ", groupID=" & groupID & ", VarType(TileViewColumns)=" & VarType(TileViewColumns) & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", OverlayIndex=" & OverlayIndex & ", StateImageIndex=" & StateImageIndex & ", itemStates=" & itemStates & ", dontAskAgain=" & dontAskAgain
  AddLogEntry str

  If requestedInfo And RequestedInfoConstants.riItemText Then
    If listSubItem Is Nothing Then
      itemText = "Item " & (listItem.Index + 1)
    ElseIf listSubItem.Index = 0 Then
      itemText = "Item " & (listItem.Index + 1)
    Else
      itemText = "Item " & (listItem.Index + 1) & ", SubItem " & listSubItem.Index
    End If
  End If
  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = 0
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

Private Sub ExLVwA_ItemGetGroup(ByVal itemIndex As Long, ByVal occurrenceIndex As Long, groupIndex As Long)
  AddLogEntry "ExLVwA_ItemGetGroup: itemIndex=" & itemIndex & ", occurrenceIndex=" & occurrenceIndex & ", groupIndex=" & groupIndex
End Sub

Private Sub ExLVwA_ItemGetInfoTipText(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemGetInfoTipText: listItem=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "ExLVwA_ItemGetInfoTipText: listItem=" & listItem.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If ExLVwA.VirtualMode Then
      infoTipText = "Hello world!"
    Else
      If infoTipText <> "" Then
        infoTipText = infoTipText & vbNewLine & "ID: " & listItem.ID & vbNewLine & "ItemData: " & listItem.ItemData
      Else
        infoTipText = "ID: " & listItem.ID & vbNewLine & "ItemData: " & listItem.ItemData
      End If
    End If
  End If
End Sub

Private Sub ExLVwA_ItemGetOccurrencesCount(ByVal itemIndex As Long, occurrencesCount As Long)
  AddLogEntry "ExLVwA_ItemGetOccurrencesCount: itemIndex=" & itemIndex & ", occurrencesCount=" & occurrencesCount
End Sub

Private Sub ExLVwA_ItemMouseEnter(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemMouseEnter: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_ItemMouseEnter: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_ItemMouseLeave(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemMouseLeave: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_ItemMouseLeave: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_ItemSelectionChanged(ByVal listItem As ExLVwLibACtl.IListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemSelectionChanged: listItem=Nothing"
  Else
    AddLogEntry "ExLVwA_ItemSelectionChanged: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwA_ItemSetText(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal itemText As String)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemSetText: listItem=Nothing, itemText=" & itemText
  Else
    AddLogEntry "ExLVwA_ItemSetText: listItem=" & listItem.Index & ", itemText=" & itemText
  End If
End Sub

Private Sub ExLVwA_ItemStateImageChanged(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExLVwLibACtl.StateImageChangeCausedByConstants)
  Dim li As ExLVwLibACtl.ListViewItem
  Dim overallStateImageIndex As Long

  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemStateImageChanged: listItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  Else
    AddLogEntry "ExLVwA_ItemStateImageChanged: listItem=" & listItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy

    If bComctl32Version610OrNewer And ExLVwA.ShowHeaderStateImages Then
      overallStateImageIndex = 2
      For Each li In ExLVwA.ListItems
        If li.StateImageIndex = 1 Then
          overallStateImageIndex = 1
          Exit For
        End If
      Next li
      ExLVwA.Columns(0).StateImageIndex = overallStateImageIndex
    End If
  End If
End Sub

Private Sub ExLVwA_ItemStateImageChanging(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExLVwLibACtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_ItemStateImageChanging: listItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  Else
    AddLogEntry "ExLVwA_ItemStateImageChanging: listItem=" & listItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExLVwA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
  If keyCode = KeyCodeConstants.vbKeyF2 Then
    ExLVwA.CaretItem.StartLabelEditing
  End If
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (ExLVwA.DraggedItems Is Nothing) Then ExLVwA.EndDrag True
  End If
End Sub

Private Sub ExLVwA_KeyPress(keyAscii As Integer)
  AddLogEntry "ExLVwA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExLVwA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwA_LostFocus()
  AddLogEntry "ExLVwA_LostFocus"
End Sub

Private Sub ExLVwA_MapGroupWideToTotalItemIndex(ByVal groupIndex As Long, ByVal groupWideItemIndex As Long, totalItemIndex As Long)
  AddLogEntry "ExLVwA_MapGroupWideToTotalItemIndex: groupIndex=" & groupIndex & ", groupWideItemIndex=" & groupWideItemIndex & ", totalItemIndex=" & totalItemIndex
End Sub

Private Sub ExLVwA_MClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_MDblClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_MouseDown(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MouseDown: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_MouseEnter(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MouseEnter: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_MouseHover(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MouseHover: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_MouseLeave(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MouseLeave: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_MouseMove(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MouseMove: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub ExLVwA_MouseUp(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_MouseUp: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If button = MouseButtonConstants.vbLeftButton Then
    If Not (ExLVwA.DraggedItems Is Nothing) Then
      ' are we within the client area?
      If ((HitTestConstants.htAbove Or HitTestConstants.htBelow Or HitTestConstants.htToLeft Or HitTestConstants.htToRight) And hitTestDetails) = 0 Then
        ' yes
        ExLVwA.EndDrag False
      Else
        ' no
        ExLVwA.EndDrag True
      End If
    End If
  End If
End Sub

Private Sub ExLVwA_MouseWheel(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants, ByVal scrollAxis As ExLVwLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  Dim str As String

  str = "ExLVwA_MouseWheel: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta

  AddLogEntry str
End Sub

Private Sub ExLVwA_OLECompleteDrag(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal performedEffect As ExLVwLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwA_OLEDragDrop(ByVal data As ExLVwLibACtl.IOLEDataObject, effect As ExLVwLibACtl.OLEDropEffectConstants, dropTarget As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFDIB) Then
    Set draggedPictureA = data.GetData(vbCFDIB)
    ExLVwA.BkImage = draggedPictureA.Handle
  End If
  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ExLVwA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwA_OLEDragEnter(ByVal data As ExLVwLibACtl.IOLEDataObject, effect As ExLVwLibACtl.OLEDropEffectConstants, dropTarget As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ExLVwA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ExLVwA_OLEDragLeave(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal dropTarget As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_OLEDragLeavePotentialTarget()
  AddLogEntry "ExLVwA_OLEDragLeavePotentialTarget"
End Sub

Private Sub ExLVwA_OLEDragMouseMove(ByVal data As ExLVwLibACtl.IOLEDataObject, effect As ExLVwLibACtl.OLEDropEffectConstants, dropTarget As ExLVwLibACtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwA_OLEGiveFeedback(ByVal effect As ExLVwLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ExLVwA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ExLVwA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As ExLVwLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "ExLVwA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ExLVwA_OLEReceivedNewData(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwA_OLESetData(ByVal data As ExLVwLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwA_OLEStartDrag(ByVal data As ExLVwLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ExLVwA_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwA_OwnerDrawItem(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal itemState As ExLVwLibACtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibACtl.RECTANGLE)
  Dim str As String

  str = "ExLVwA_OwnerDrawItem: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub ExLVwA_RClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_RClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_RDblClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_RDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExLVwA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertGroupsA
  InsertItemsA
  If bComctl32Version610OrNewer Then InsertFooterItemsA
End Sub

Private Sub ExLVwA_RemovedColumn(ByVal column As ExLVwLibACtl.IVirtualListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_RemovedColumn: column=Nothing"
  Else
    AddLogEntry "ExLVwA_RemovedColumn: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwA_RemovedGroup(ByVal Group As ExLVwLibACtl.IVirtualListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_RemovedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwA_RemovedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwA_RemovedItem(ByVal listItem As ExLVwLibACtl.IVirtualListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_RemovedItem: listItem=Nothing"
  Else
    AddLogEntry "ExLVwA_RemovedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwA_RemovingColumn(ByVal column As ExLVwLibACtl.IListViewColumn, cancelDeletion As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_RemovingColumn: column=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExLVwA_RemovingColumn: column=" & column.Caption & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExLVwA_RemovingGroup(ByVal Group As ExLVwLibACtl.IListViewGroup, cancelDeletion As Boolean)
  If Group Is Nothing Then
    AddLogEntry "ExLVwA_RemovingGroup: group=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExLVwA_RemovingGroup: group=" & Group.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExLVwA_RemovingItem(ByVal listItem As ExLVwLibACtl.IListViewItem, cancelDeletion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_RemovingItem: listItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExLVwA_RemovingItem: listItem=" & listItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExLVwA_RenamedItem(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal previousItemText As String, ByVal newItemText As String)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_RenamedItem: listItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText
  Else
    AddLogEntry "ExLVwA_RenamedItem: listItem=" & listItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText
  End If
End Sub

Private Sub ExLVwA_RenamingItem(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal previousItemText As String, ByVal newItemText As String, cancelRenaming As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_RenamingItem: listItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  Else
    AddLogEntry "ExLVwA_RenamingItem: listItem=" & listItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  End If
End Sub

Private Sub ExLVwA_ResizedControlWindow()
  AddLogEntry "ExLVwA_ResizedControlWindow"
End Sub

Private Sub ExLVwA_ResizingColumn(ByVal column As ExLVwLibACtl.IListViewColumn, newColumnWidth As Long, Cancel As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwA_ResizingColumn: column=Nothing, newColumnWidth=" & newColumnWidth & ", cancel=" & Cancel
  Else
    AddLogEntry "ExLVwA_ResizingColumn: column=" & column.Caption & ", newColumnWidth=" & newColumnWidth & ", cancel=" & Cancel
  End If
End Sub

Private Sub ExLVwA_SelectedItemRange(ByVal firstItem As ExLVwLibACtl.IListViewItem, ByVal lastItem As ExLVwLibACtl.IListViewItem)
  Dim str As String

  str = "ExLVwA_SelectedItemRange: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, lastItem="
  Else
    str = str & firstItem.Text & ", lastItem="
  End If
  If lastItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & lastItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwA_SettingItemInfoTipText(ByVal listItem As ExLVwLibACtl.IListViewItem, infoTipText As String, abortInfoTip As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_SettingItemInfoTipText: listItem=Nothing, infoTipText=" & infoTipText & ", abortInfoTip=" & abortInfoTip
  Else
    AddLogEntry "ExLVwA_SettingItemInfoTipText: listItem=" & listItem.Text & ", infoTipText=" & infoTipText & ", abortInfoTip=" & abortInfoTip
  End If
End Sub

Private Sub ExLVwA_StartingLabelEditing(ByVal listItem As ExLVwLibACtl.IListViewItem, cancelEditing As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwA_StartingLabelEditing: listItem=Nothing, cancelEditing=" & cancelEditing
  Else
    AddLogEntry "ExLVwA_StartingLabelEditing: listItem=" & listItem.Text & ", cancelEditing=" & cancelEditing
  End If
End Sub

Private Sub ExLVwA_SubItemMouseEnter(ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwA_SubItemMouseEnter: listSubItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_SubItemMouseEnter: listSubItem=" & listSubItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_SubItemMouseLeave(ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwA_SubItemMouseLeave: listSubItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwA_SubItemMouseLeave: listSubItem=" & listSubItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwA_Validate(Cancel As Boolean)
  AddLogEntry "ExLVwA_Validate"
End Sub

Private Sub ExLVwA_XClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_XClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwA_XDblClick(ByVal listItem As ExLVwLibACtl.IListViewItem, ByVal listSubItem As ExLVwLibACtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExLVwA_XDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_AbortedDrag()
  AddLogEntry "ExLVwU_AbortedDrag"
  Set ExLVwU.DropHilitedItem = Nothing
End Sub

Private Sub ExLVwU_AfterScroll(ByVal dx As Long, ByVal dy As Long)
  AddLogEntry "ExLVwU_AfterScroll: dx=" & dx & ", dy=" & dy
End Sub

Private Sub ExLVwU_BeforeScroll(ByVal dx As Long, ByVal dy As Long)
  AddLogEntry "ExLVwU_BeforeScroll: dx=" & dx & ", dy=" & dy
End Sub

Private Sub ExLVwU_BeginColumnResizing(ByVal column As ExLVwLibUCtl.IListViewColumn, Cancel As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_BeginColumnResizing: column=Nothing, cancel=" & Cancel
  Else
    AddLogEntry "ExLVwU_BeginColumnResizing: column=" & column.Caption & ", cancel=" & Cancel
  End If
End Sub

Private Sub ExLVwU_BeginMarqueeSelection(Cancel As Boolean)
  AddLogEntry "ExLVwU_BeginMarqueeSelection: cancel=" & Cancel
End Sub

Private Sub ExLVwU_CacheItemsHint(ByVal FirstItem As ExLVwLibUCtl.IListViewItem, ByVal lastItem As ExLVwLibUCtl.IListViewItem)
  AddLogEntry "ExLVwU_CacheItemsHint: firstItem=" & FirstItem.Index & ", lastItem=" & lastItem.Index
End Sub

Private Sub ExLVwU_CancelSubItemEdit(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants)
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwU_CancelSubItemEdit: listItem=Nothing, listSubItem=Nothing, editMode=" & editMode
  Else
    AddLogEntry "ExLVwU_CancelSubItemEdit: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", editMode=" & editMode
  End If
End Sub

Private Sub ExLVwU_CaretChanged(ByVal previousCaretItem As ExLVwLibUCtl.IListViewItem, ByVal newCaretItem As ExLVwLibUCtl.IListViewItem)
  Dim str As String

  str = "ExLVwU_CaretChanged: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwU_ChangedSortOrder(ByVal previousSortOrder As ExLVwLibUCtl.SortOrderConstants, ByVal newSortOrder As ExLVwLibUCtl.SortOrderConstants)
  AddLogEntry "ExLVwU_ChangedSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder
End Sub

Private Sub ExLVwU_ChangedWorkAreas(ByVal WorkAreas As ExLVwLibUCtl.IListViewWorkAreas)
  If WorkAreas Is Nothing Then
    AddLogEntry "ExLVwU_ChangedWorkAreas: WorkAreas=Nothing"
  Else
    AddLogEntry "ExLVwU_ChangedWorkAreas: WorkAreas=" & WorkAreas.Count
  End If
End Sub

Private Sub ExLVwU_ChangingSortOrder(ByVal previousSortOrder As ExLVwLibUCtl.SortOrderConstants, ByVal newSortOrder As ExLVwLibUCtl.SortOrderConstants, cancelChange As Boolean)
  AddLogEntry "ExLVwU_ChangingSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder & ", cancelChange=" & cancelChange
End Sub

Private Sub ExLVwU_ChangingWorkAreas(ByVal WorkAreas As ExLVwLibUCtl.IVirtualListViewWorkAreas, cancelChanges As Boolean)
  If WorkAreas Is Nothing Then
    AddLogEntry "ExLVwU_ChangingWorkAreas: WorkAreas=Nothing, cancelChanges=" & cancelChanges
  Else
    AddLogEntry "ExLVwU_ChangingWorkAreas: WorkAreas=" & WorkAreas.Count & ", cancelChanges=" & cancelChanges
  End If
End Sub

Private Sub ExLVwU_Click(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_Click: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_CollapsedGroup(ByVal Group As ExLVwLibUCtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_CollapsedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwU_CollapsedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_ColumnBeginDrag(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, doAutomaticDragDrop As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnBeginDrag: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDragDrop=" & doAutomaticDragDrop
  Else
    AddLogEntry "ExLVwU_ColumnBeginDrag: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDragDrop=" & doAutomaticDragDrop
  End If
End Sub

Private Sub ExLVwU_ColumnClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  Dim col As ExLVwLibUCtl.ListViewColumn
  Dim curArrow As ExLVwLibUCtl.SortArrowConstants
  Dim newArrow As ExLVwLibUCtl.SortArrowConstants

  If column Is Nothing Then
    ' shouldn't happen
    AddLogEntry "ExLVwU_ColumnClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_ColumnClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

    ' update the sort arrows
    For Each col In ExLVwU.Columns
      If col.Index = column.Index Then
        If bComctl32Version600OrNewer Then
          curArrow = col.SortArrow
        Else
          Select Case col.BitmapHandle
            Case 0
              curArrow = saNone
            Case hBMPDownArrow
              curArrow = saDown
            Case hBMPUpArrow
              curArrow = saUp
          End Select
        End If

        If curArrow = saUp Then
          newArrow = saDown
        Else
          newArrow = saUp
        End If
      Else
        newArrow = saNone
      End If

      If bComctl32Version600OrNewer Then
        col.SortArrow = newArrow
        Select Case newArrow
          Case saDown
            ExLVwU.SortOrder = soDescending
            Set ExLVwU.SelectedColumn = col
            ExLVwU.SortItems , , , , , col, True
          Case saUp
            ExLVwU.SortOrder = soAscending
            Set ExLVwU.SelectedColumn = col
            ExLVwU.SortItems , , , , , col, True
        End Select
      Else
        Select Case newArrow
          Case saNone
            col.BitmapHandle = 0
          Case saDown
            col.BitmapHandle = hBMPDownArrow
            ExLVwU.SortOrder = soDescending
            ExLVwU.SortItems , , , , , col
          Case saUp
            col.BitmapHandle = hBMPUpArrow
            ExLVwU.SortOrder = soAscending
            ExLVwU.SortItems , , , , , col
        End Select
      End If
    Next col
  End If
End Sub

Private Sub ExLVwU_ColumnDropDown(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnDropDown: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExLVwU_ColumnDropDown: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu

    ' TODO: Display drop-down
  End If
End Sub

Private Sub ExLVwU_ColumnEndAutoDragDrop(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, doAutomaticDrop As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnEndAutoDragDrop: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDrop=" & doAutomaticDrop
  Else
    AddLogEntry "ExLVwU_ColumnEndAutoDragDrop: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", doAutomaticDrop=" & doAutomaticDrop
  End If
End Sub

Private Sub ExLVwU_ColumnMouseEnter(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnMouseEnter: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_ColumnMouseEnter: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_ColumnMouseLeave(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnMouseLeave: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_ColumnMouseLeave: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_ColumnStateImageChanged(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExLVwLibUCtl.StateImageChangeCausedByConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnStateImageChanged: column=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  Else
    AddLogEntry "ExLVwU_ColumnStateImageChanged: column=" & column.Caption & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy

    If column.Index = 0 Then
      ExLVwU.ListItems(-1).StateImageIndex = column.StateImageIndex
    End If
  End If
End Sub

Private Sub ExLVwU_ColumnStateImageChanging(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExLVwLibUCtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ColumnStateImageChanging: column=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  Else
    AddLogEntry "ExLVwU_ColumnStateImageChanging: column=" & column.Caption & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExLVwU_CompareGroups(ByVal firstGroup As ExLVwLibUCtl.IListViewGroup, ByVal secondGroup As ExLVwLibUCtl.IListViewGroup, result As ExLVwLibUCtl.CompareResultConstants)
  Dim str As String

  str = "ExLVwU_CompareGroups: firstGroup="
  If firstGroup Is Nothing Then
    str = str & "Nothing, secondGroup="
  Else
    str = str & firstGroup.Text & ", secondGroup="
  End If
  If secondGroup Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondGroup.Text
  End If
  str = str & ", result=" & result

  AddLogEntry str
End Sub

Private Sub ExLVwU_CompareItems(ByVal firstItem As ExLVwLibUCtl.IListViewItem, ByVal secondItem As ExLVwLibUCtl.IListViewItem, result As ExLVwLibUCtl.CompareResultConstants)
  Dim str As String

  str = "ExLVwU_CompareItems: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & firstItem.Text & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Text
  End If
  str = str & ", result=" & result

  AddLogEntry str
End Sub

Private Sub ExLVwU_ConfigureSubItemControl(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal controlKind As ExLVwLibUCtl.SubItemControlKindConstants, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants, ByVal subItemControl As ExLVwLibUCtl.SubItemControlConstants, themeAppName As String, themeIDList As String, hFont As Long, textColor As stdole.OLE_COLOR, pPropertyDescription As Long, ByVal pPropertyValue As Long)
  Dim buffer As String

  buffer = String$(1024, 0)
  PropVariantToString pPropertyValue, StrPtr(buffer), 1024
  buffer = Left$(buffer, lstrlen(StrPtr(buffer)))
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwU_ConfigureSubItemControl: listItem=Nothing, listSubItem=Nothing, controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl & ", themeAppName=" & themeAppName & ", themeIDList=" & themeIDList & ", hFont=0x" & Hex(hFont) & ", textColor=0x" & Hex(textColor) & ", pPropertyDescription=0x" & Hex(pPropertyDescription) & ", propertyValue=" & buffer
  Else
    AddLogEntry "ExLVwU_ConfigureSubItemControl: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl & ", themeAppName=" & themeAppName & ", themeIDList=" & themeIDList & ", hFont=0x" & Hex(hFont) & ", textColor=0x" & Hex(textColor) & ", pPropertyDescription=0x" & Hex(pPropertyDescription) & ", propertyValue=" & buffer
  End If
End Sub

Private Sub ExLVwU_ContextMenu(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  Dim str As String

  str = "ExLVwU_ContextMenu: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu

  AddLogEntry str

  If listItem Is Nothing Then
    PopupMenu mnuView, , ExLVwU.Left + ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ExLVwU.Top + ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  Else
    '
  End If
End Sub

Private Sub ExLVwU_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExLVwU_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExLVwU_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  AddLogEntry "ExLVwU_CreatedHeaderControlWindow: hWndHeader=0x" & Hex(hWndHeader)
  InsertColumnsU
End Sub

Private Sub ExLVwU_CustomDraw(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal drawAllItems As Boolean, textColor As stdole.OLE_COLOR, TextBackColor As stdole.OLE_COLOR, ByVal drawStage As ExLVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExLVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE, furtherProcessing As ExLVwLibUCtl.CustomDrawReturnValuesConstants)
'  Dim str As String
'
'  str = "ExLVwU_CustomDraw: listItem="
'  If listItem Is Nothing Then
'    str = str & "Nothing, listSubItem="
'  Else
'    str = str & listItem.Text & ", listSubItem="
'  End If
'  If listSubItem Is Nothing Then
'    str = str & "Nothing"
'  Else
'    str = str & listSubItem.Text
'  End If
'  str = str & ", drawAllItems=" & drawAllItems & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(TextBackColor) & ", drawStage=" & drawStage & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'
'  AddLogEntry str
End Sub

Private Sub ExLVwU_DblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_DblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExLVwU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ExLVwU_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExLVwU_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExLVwU_DestroyedHeaderControlWindow(ByVal hWndHeader As Long)
  AddLogEntry "ExLVwU_DestroyedHeaderControlWindow: hWndHeader=0x" & Hex(hWndHeader)
End Sub

Private Sub ExLVwU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ExLVwU_DragDrop"
End Sub

Private Sub ExLVwU_DragMouseMove(dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwU_DragMouseMove: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  Else
    AddLogEntry "ExLVwU_DragMouseMove: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  End If
End Sub

Private Sub ExLVwU_DragOver(Source As Control, x As Single, y As Single, state As Integer)
  AddLogEntry "ExLVwU_DragOver"
End Sub

Private Sub ExLVwU_Drop(ByVal dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim dx As Long
  Dim dy As Long
  Dim itm As ExLVwLibUCtl.ListViewItem
  Dim xPos As Long
  Dim yPos As Long

  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwU_Drop: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_Drop: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If

  x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  dx = x - ptStartDrag.x
  dy = y - ptStartDrag.y

  For Each itm In ExLVwU.DraggedItems
    itm.GetPosition xPos, yPos
    itm.SetPosition xPos + dx, yPos + dy
  Next itm
End Sub

Private Sub ExLVwU_EditChange()
  AddLogEntry "ExLVwU_EditChange - " & ExLVwU.EditText
End Sub

Private Sub ExLVwU_EditClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  AddLogEntry "ExLVwU_EditContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub ExLVwU_EditDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditGotFocus()
  AddLogEntry "ExLVwU_EditGotFocus"
End Sub

Private Sub ExLVwU_EditKeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwU_EditKeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwU_EditKeyPress(keyAscii As Integer)
  AddLogEntry "ExLVwU_EditKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExLVwU_EditKeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwU_EditKeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwU_EditLostFocus()
  AddLogEntry "ExLVwU_EditLostFocus"
End Sub

Private Sub ExLVwU_EditMClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
'  AddLogEntry "ExLVwU_EditMouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditMouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditMouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As ExLVwLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "ExLVwU_EditMouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub ExLVwU_EditRClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditRClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditRDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditRDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditXClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditXClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EditXDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExLVwU_EditXDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExLVwU_EmptyMarkupTextLinkClick(ByVal linkIndex As Long, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  AddLogEntry "ExLVwU_EmptyMarkupTextLinkClick: linkIndex=" & linkIndex & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub ExLVwU_EndColumnResizing(ByVal column As ExLVwLibUCtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_EndColumnResizing: column=Nothing"
  Else
    AddLogEntry "ExLVwU_EndColumnResizing: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwU_EndSubItemEdit(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants, ByVal pPropertyKey As Long, ByVal pPropertyValue As Long, Cancel As Boolean)
  Dim propKey As String
  Dim propValue As String

  propKey = String$(1024, 0)
  PSStringFromPropertyKey pPropertyKey, StrPtr(propKey), 1024
  propKey = Left$(propKey, lstrlen(StrPtr(propKey)))
  propValue = String$(1024, 0)
  PropVariantToString pPropertyValue, StrPtr(propValue), 1024
  propValue = Left$(propValue, lstrlen(StrPtr(propValue)))
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwU_EndSubItemEdit: listItem=Nothing, listSubItem=Nothing, editMode=" & editMode & ", pPropertyKey=" & propKey & ", propertyValue=" & propValue & ", Cancel=" & Cancel
  Else
    AddLogEntry "ExLVwU_EndSubItemEdit: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", editMode=" & editMode & ", pPropertyKey=" & propKey & ", propertyValue=" & propValue & ", Cancel=" & Cancel
  End If
End Sub

Private Sub ExLVwU_ExpandedGroup(ByVal Group As ExLVwLibUCtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_ExpandedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwU_ExpandedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_FilterButtonClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, filterButtonRectangle As ExLVwLibUCtl.RECTANGLE, raiseFilterChanged As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_FilterButtonClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", filterButtonRectangle=(" & filterButtonRectangle.Left & "," & filterButtonRectangle.Top & ")-(" & filterButtonRectangle.Right & "," & filterButtonRectangle.Bottom & "), raiseFilterChanged=" & raiseFilterChanged
  Else
    AddLogEntry "ExLVwU_FilterButtonClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", filterButtonRectangle=(" & filterButtonRectangle.Left & "," & filterButtonRectangle.Top & ")-(" & filterButtonRectangle.Right & "," & filterButtonRectangle.Bottom & "), raiseFilterChanged=" & raiseFilterChanged
  End If
End Sub

Private Sub ExLVwU_FilterChanged(ByVal column As ExLVwLibUCtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_FilterChanged: column=Nothing"
  Else
    AddLogEntry "ExLVwU_FilterChanged: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwU_FindVirtualItem(ByVal itemToStartWith As ExLVwLibUCtl.IListViewItem, ByVal searchMode As ExLVwLibUCtl.SearchModeConstants, searchFor As Variant, ByVal searchDirection As ExLVwLibUCtl.SearchDirectionConstants, ByVal wrapAtLastItem As Boolean, foundItem As ExLVwLibUCtl.IListViewItem)
  Dim str As String

  str = "ExLVwU_FindVirtualItem: itemToStartWith="
  If itemToStartWith Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & itemToStartWith.Text
  End If
  str = str & ", searchMode=" & searchMode & ", VarType(searchFor)=" & VarType(searchFor) & ", searchDirection=" & searchDirection & ", wrapAtLastItem=" & wrapAtLastItem & ", foundItem="
  If foundItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & foundItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwU_FooterItemClick(ByVal footerItem As ExLVwLibUCtl.IListViewFooterItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, removeFooterArea As Boolean)
  If footerItem Is Nothing Then
    AddLogEntry "ExLVwU_FooterItemClick: footerItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", removeFooterArea=" & removeFooterArea
  Else
    AddLogEntry "ExLVwU_FooterItemClick: footerItem=" & footerItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", removeFooterArea=" & removeFooterArea

    Select Case footerItem.ItemData
      Case 1
        MsgBox "Good night!"
      Case 2
        MsgBox "Let's go, Pinky!"
    End Select
  End If
End Sub

Private Sub ExLVwU_FreeColumnData(ByVal column As ExLVwLibUCtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_FreeColumnData: column=Nothing"
  Else
    AddLogEntry "ExLVwU_FreeColumnData: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwU_FreeFooterItemData(ByVal footerItem As ExLVwLibUCtl.IListViewFooterItem, ByVal ItemData As Long)
  If footerItem Is Nothing Then
    AddLogEntry "ExLVwU_FreeFooterItemData: footerItem=Nothing, itemData=" & ItemData
  Else
    AddLogEntry "ExLVwU_FreeFooterItemData: footerItem=" & footerItem.Text & ", itemData=" & ItemData
  End If
End Sub

Private Sub ExLVwU_FreeItemData(ByVal listItem As ExLVwLibUCtl.IListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_FreeItemData: listItem=Nothing"
  Else
    AddLogEntry "ExLVwU_FreeItemData: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwU_GetSubItemControl(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal controlKind As ExLVwLibUCtl.SubItemControlKindConstants, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants, subItemControl As ExLVwLibUCtl.SubItemControlConstants)
  If listSubItem Is Nothing Then
'    AddLogEntry "ExLVwU_GetSubItemControl: listItem=Nothing, listSubItem=Nothing, controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl
  Else
'    AddLogEntry "ExLVwU_GetSubItemControl: listItem=" & listSubItem.ParentItem.Text & ", listSubItem=" & listSubItem.Text & ", controlKind=" & controlKind & ", editMode=" & editMode & ", subItemControl=" & subItemControl
  End If
End Sub

Private Sub ExLVwU_GotFocus()
  AddLogEntry "ExLVwU_GotFocus"
  Set objActiveCtl = ExLVwU
  UpdateMenu
End Sub

Private Sub ExLVwU_GroupAsynchronousDrawFailed(ByVal Group As ExLVwLibUCtl.IListViewGroup, imageDetails As ExLVwLibUCtl.FAILEDIMAGEDETAILS, ByVal failureReason As ExLVwLibUCtl.ImageDrawingFailureReasonConstants, furtherProcessing As ExLVwLibUCtl.FailedAsyncDrawReturnValuesConstants, newImageToDraw As Long)
  Dim str As String

  str = "ExLVwU_GroupAsynchronousDrawFailed: group="
  If Group Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Group.Text
  End If
  str = str & ", failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  AddLogEntry "     imageDetails.hImageList=0x" & Hex(imageDetails.hImageList)
  AddLogEntry "     imageDetails.hDC=0x" & Hex(imageDetails.hDC)
  AddLogEntry "     imageDetails.IconIndex=" & imageDetails.IconIndex
  AddLogEntry "     imageDetails.OverlayIndex=" & imageDetails.OverlayIndex
  AddLogEntry "     imageDetails.DrawingStyle=" & imageDetails.DrawingStyle
  AddLogEntry "     imageDetails.DrawingEffects=" & imageDetails.DrawingEffects
  AddLogEntry "     imageDetails.BackColor=0x" & Hex(imageDetails.BackColor)
  AddLogEntry "     imageDetails.ForeColor=0x" & Hex(imageDetails.ForeColor)
  AddLogEntry "     imageDetails.EffectColor=0x" & Hex(imageDetails.EffectColor)
End Sub

Private Sub ExLVwU_GroupCustomDraw(ByVal Group As ExLVwLibUCtl.IListViewGroup, textColor As stdole.OLE_COLOR, headerAlignment As ExLVwLibUCtl.AlignmentConstants, footerAlignment As ExLVwLibUCtl.AlignmentConstants, ByVal drawStage As ExLVwLibUCtl.CustomDrawStageConstants, ByVal groupState As ExLVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE, textRectangle As ExLVwLibUCtl.RECTANGLE, furtherProcessing As ExLVwLibUCtl.CustomDrawReturnValuesConstants)
'  If Group Is Nothing Then
'    AddLogEntry "ExLVwU_GroupCustomDraw: Group=Nothing, textColor=0x" & Hex(textColor) & ", headerAlignment=" & headerAlignment & ", footerAlignment=" & footerAlignment & ", drawStage=" & drawStage & ", groupState=" & groupState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ExLVwU_GroupCustomDraw: Group=" & Group.Text & ", textColor=0x" & Hex(textColor) & ", headerAlignment=" & headerAlignment & ", footerAlignment=" & footerAlignment & ", drawStage=" & drawStage & ", groupState=" & groupState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ExLVwU_GroupGotFocus(ByVal Group As ExLVwLibUCtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_GroupGotFocus: group=Nothing"
  Else
    AddLogEntry "ExLVwU_GroupGotFocus: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_GroupLostFocus(ByVal Group As ExLVwLibUCtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_GroupLostFocus: group=Nothing"
  Else
    AddLogEntry "ExLVwU_GroupLostFocus: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_GroupSelectionChanged(ByVal Group As ExLVwLibUCtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_GroupSelectionChanged: group=Nothing"
  Else
    AddLogEntry "ExLVwU_GroupSelectionChanged: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_GroupTaskLinkClick(ByVal Group As ExLVwLibUCtl.IListViewGroup, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_GroupTaskLinkClick: group=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_GroupTaskLinkClick: group=" & Group.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderAbortedDrag()
  AddLogEntry "ExLVwU_HeaderAbortedDrag"
End Sub

Private Sub ExLVwU_HeaderChevronClick(ByVal firstOverflownColumn As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, showDefaultMenu As Boolean)
  If firstOverflownColumn Is Nothing Then
    AddLogEntry "ExLVwU_HeaderChevronClick: firstOverflownColumn=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExLVwU_HeaderChevronClick: firstOverflownColumn=" & firstOverflownColumn.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", showDefaultMenu=" & showDefaultMenu
  End If
End Sub

Private Sub ExLVwU_HeaderClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderContextMenu(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, showDefaultMenu As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderContextMenu: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExLVwU_HeaderContextMenu: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", showDefaultMenu=" & showDefaultMenu
  End If
End Sub

Private Sub ExLVwU_HeaderCustomDraw(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal drawStage As ExLVwLibUCtl.CustomDrawStageConstants, ByVal columnState As ExLVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE, furtherProcessing As ExLVwLibUCtl.CustomDrawReturnValuesConstants)
'  If column Is Nothing Then
'    AddLogEntry "ExLVwU_HeaderCustomDraw: column=Nothing, drawStage=" & drawStage & ", columnState=" & columnState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ExLVwU_HeaderCustomDraw: column=" & column.Caption & ", drawStage=" & drawStage & ", columnState=" & columnState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ExLVwU_HeaderDblClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderDividerDblClick(ByVal column As ExLVwLibUCtl.IListViewColumn, autoSizeColumn As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderDividerDblClick: column=Nothing, autoSizeColumn=" & autoSizeColumn
  Else
    AddLogEntry "ExLVwU_HeaderDividerDblClick: column=" & column.Caption & ", autoSizeColumn=" & autoSizeColumn
  End If
End Sub

Private Sub ExLVwU_HeaderDragMouseMove(dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwU_HeaderDragMouseMove: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  Else
    AddLogEntry "ExLVwU_HeaderDragMouseMove: dropTarget=" & dropTarget.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  End If
End Sub

Private Sub ExLVwU_HeaderDrop(ByVal dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If dropTarget Is Nothing Then
    AddLogEntry "ExLVwU_HeaderDrop: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderDrop: dropTarget=" & dropTarget.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderGotFocus()
  AddLogEntry "ExLVwU_HeaderGotFocus"
End Sub

Private Sub ExLVwU_HeaderItemGetDisplayInfo(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal requestedInfo As ExLVwLibUCtl.RequestedInfoConstants, IconIndex As Long, ByVal maxColumnCaptionLength As Long, columnCaption As String, dontAskAgain As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderItemGetDisplayInfo: column=Nothing, requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", maxColumnCaptionLength=" & maxColumnCaptionLength & ", columnCaption=" & columnCaption & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "ExLVwU_HeaderItemGetDisplayInfo: column=" & column.Index & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", maxColumnCaptionLength=" & maxColumnCaptionLength & ", columnCaption=" & columnCaption & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub ExLVwU_HeaderKeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwU_HeaderKeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwU_HeaderKeyPress(keyAscii As Integer)
  AddLogEntry "ExLVwU_HeaderKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExLVwU_HeaderKeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwU_HeaderKeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwU_HeaderLostFocus()
  AddLogEntry "ExLVwU_HeaderLostFocus"
End Sub

Private Sub ExLVwU_HeaderMClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMDblClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseDown(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMouseDown: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMouseDown: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseEnter(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMouseEnter: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMouseEnter: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseHover(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMouseHover: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMouseHover: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseLeave(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMouseLeave: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMouseLeave: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseMove(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
'    AddLogEntry "ExLVwU_HeaderMouseMove: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
'    AddLogEntry "ExLVwU_HeaderMouseMove: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseUp(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMouseUp: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderMouseUp: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderMouseWheel(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, ByVal scrollAxis As ExLVwLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderMouseWheel: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
  Else
    AddLogEntry "ExLVwU_HeaderMouseWheel: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
  End If
End Sub

Private Sub ExLVwU_HeaderOLECompleteDrag(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal performedEffect As ExLVwLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwU_HeaderOLEDragDrop(ByVal data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, ByVal dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFDIB) Then
    Set draggedPictureU = data.GetData(vbCFDIB)
    ExLVwU.BkImage = draggedPictureU.Handle
  End If
  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ExLVwU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwU_HeaderOLEDragEnter(ByVal data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderOLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ExLVwU_HeaderOLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ExLVwU_HeaderOLEDragLeave(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderOLEDragLeavePotentialTarget()
  AddLogEntry "ExLVwU_HeaderOLEDragLeavePotentialTarget"
End Sub

Private Sub ExLVwU_HeaderOLEDragMouseMove(ByVal data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Caption
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", xListView=" & xListView & ", yListView=" & yListView & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderOLEGiveFeedback(ByVal effect As ExLVwLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ExLVwU_HeaderOLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ExLVwU_HeaderOLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As ExLVwLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "ExLVwU_HeaderOLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ExLVwU_HeaderOLEReceivedNewData(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderOLESetData(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderOLEStartDrag(ByVal data As ExLVwLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_HeaderOLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderOwnerDrawItem(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal itemState As ExLVwLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE)
  Dim str As String

  str = "ExLVwU_HeaderOwnerDrawItem: column="
  If column Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & column.Caption
  End If
  str = str & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub ExLVwU_HeaderRClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderRClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderRClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderRDblClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderRDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderRDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderXClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderXClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderXClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HeaderXDblClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_HeaderXDblClick: column=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_HeaderXDblClick: column=" & column.Caption & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_HotItemChanged(ByVal previousHotItem As ExLVwLibUCtl.IListViewItem, ByVal newHotItem As ExLVwLibUCtl.IListViewItem)
  Dim str As String

  str = "ExLVwU_HotItemChanged: previousHotItem="
  If previousHotItem Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotItem.Text & ", newHotItem="
  End If
  If newHotItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwU_HotItemChanging(ByVal previousHotItem As ExLVwLibUCtl.IListViewItem, ByVal newHotItem As ExLVwLibUCtl.IListViewItem, cancelChange As Boolean)
  Dim str As String

  str = "ExLVwU_HotItemChanging: previousHotItem="
  If previousHotItem Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotItem.Text & ", newHotItem="
  End If
  If newHotItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotItem.Text
  End If
  str = str & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ExLVwU_IncrementalSearching(ByVal currentSearchString As String, itemToSelect As Long)
  AddLogEntry "ExLVwU_IncrementalSearching: currentSearchString=" & currentSearchString & ", itemToSelect=" & itemToSelect
End Sub

Private Sub ExLVwU_IncrementalSearchStringChanging(ByVal currentSearchString As String, ByVal keyCodeOfCharToBeAdded As Integer, cancelChange As Boolean)
  AddLogEntry "ExLVwU_IncrementalSearchStringChanging: currentSearchString=" & currentSearchString & ", keyCodeOfCharToBeAdded=" & keyCodeOfCharToBeAdded & ", cancelChange=" & cancelChange
End Sub

Private Sub ExLVwU_InsertedColumn(ByVal column As ExLVwLibUCtl.IListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_InsertedColumn: column=Nothing"
  Else
    AddLogEntry "ExLVwU_InsertedColumn: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwU_InsertedGroup(ByVal Group As ExLVwLibUCtl.IListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_InsertedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwU_InsertedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_InsertedItem(ByVal listItem As ExLVwLibUCtl.IListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_InsertedItem: listItem=Nothing"
  Else
    AddLogEntry "ExLVwU_InsertedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwU_InsertingColumn(ByVal column As ExLVwLibUCtl.IVirtualListViewColumn, cancelInsertion As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_InsertingColumn: column=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExLVwU_InsertingColumn: column=" & column.Caption & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExLVwU_InsertingGroup(ByVal Group As ExLVwLibUCtl.IVirtualListViewGroup, cancelInsertion As Boolean)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_InsertingGroup: group=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExLVwU_InsertingGroup: group=" & Group.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExLVwU_InsertingItem(ByVal listItem As ExLVwLibUCtl.IVirtualListViewItem, cancelInsertion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_InsertingItem: listItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExLVwU_InsertingItem: listItem=" & listItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExLVwU_InvokeVerbFromSubItemControl(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal verb As String)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_InvokeVerbFromSubItemControl: listItem=Nothing, verb=" & verb
  Else
    AddLogEntry "ExLVwU_InvokeVerbFromSubItemControl: listItem=" & listItem.Text & ", verb=" & verb
  End If
End Sub

Private Sub ExLVwU_ItemActivate(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal shift As Integer, ByVal x As Long, ByVal y As Long)
  Dim str As String

  str = "ExLVwU_ItemActivate: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ExLVwU_ItemAsynchronousDrawFailed(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, imageDetails As ExLVwLibUCtl.FAILEDIMAGEDETAILS, ByVal failureReason As ExLVwLibUCtl.ImageDrawingFailureReasonConstants, furtherProcessing As ExLVwLibUCtl.FailedAsyncDrawReturnValuesConstants, newImageToDraw As Long)
  Dim str As String

  str = "ExLVwU_ItemAsynchronousDrawFailed: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  AddLogEntry "     imageDetails.hImageList=0x" & Hex(imageDetails.hImageList)
  AddLogEntry "     imageDetails.hDC=0x" & Hex(imageDetails.hDC)
  AddLogEntry "     imageDetails.IconIndex=" & imageDetails.IconIndex
  AddLogEntry "     imageDetails.OverlayIndex=" & imageDetails.OverlayIndex
  AddLogEntry "     imageDetails.DrawingStyle=" & imageDetails.DrawingStyle
  AddLogEntry "     imageDetails.DrawingEffects=" & imageDetails.DrawingEffects
  AddLogEntry "     imageDetails.BackColor=0x" & Hex(imageDetails.BackColor)
  AddLogEntry "     imageDetails.ForeColor=0x" & Hex(imageDetails.ForeColor)
  AddLogEntry "     imageDetails.EffectColor=0x" & Hex(imageDetails.EffectColor)
End Sub

Private Sub ExLVwU_ItemBeginDrag(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim col As ExLVwLibUCtl.ListViewItems
  Dim str As String

  str = "ExLVwU_ItemBeginDrag: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  Set col = ExLVwU.ListItems
  col.FilterType(ExLVwLibUCtl.FilteredPropertyConstants.fpSelected) = ExLVwLibUCtl.FilterTypeConstants.ftIncluding
  col.Filter(ExLVwLibUCtl.FilteredPropertyConstants.fpSelected) = Array(True)

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ExLVwU.BeginDrag ExLVwU.CreateItemContainer(col), -1
End Sub

Private Sub ExLVwU_ItemBeginRDrag(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_ItemBeginRDrag: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_ItemGetDisplayInfo(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal requestedInfo As ExLVwLibUCtl.RequestedInfoConstants, IconIndex As Long, Indent As Long, groupID As Long, TileViewColumns() As ExLVwLibUCtl.TILEVIEWSUBITEM, ByVal maxItemTextLength As Long, itemText As String, OverlayIndex As Long, StateImageIndex As Long, itemStates As ExLVwLibUCtl.ItemStateConstants, dontAskAgain As Boolean)
  Dim c(1 To 2) As ExLVwLibUCtl.TILEVIEWSUBITEM
  Dim str As String

  str = "ExLVwU_ItemGetDisplayInfo: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Index & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Index
  End If
  str = str & ", requestedInfo=" & requestedInfo & ", IconIndex=" & IconIndex & ", Indent=" & Indent & ", groupID=" & groupID & ", VarType(TileViewColumns)=" & VarType(TileViewColumns) & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", OverlayIndex=" & OverlayIndex & ", StateImageIndex=" & StateImageIndex & ", itemStates=" & itemStates & ", dontAskAgain=" & dontAskAgain
  AddLogEntry str

  If requestedInfo And RequestedInfoConstants.riItemText Then
    If listSubItem Is Nothing Then
      itemText = "Item " & (listItem.Index + 1)
    ElseIf listSubItem.Index = 0 Then
      itemText = "Item " & (listItem.Index + 1)
    Else
      itemText = "Item " & (listItem.Index + 1) & ", SubItem " & listSubItem.Index
    End If
  End If
  If requestedInfo And RequestedInfoConstants.riIconIndex Then
    IconIndex = 0
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

Private Sub ExLVwU_ItemGetGroup(ByVal itemIndex As Long, ByVal occurrenceIndex As Long, groupIndex As Long)
  AddLogEntry "ExLVwU_ItemGetGroup: itemIndex=" & itemIndex & ", occurrenceIndex=" & occurrenceIndex & ", groupIndex=" & groupIndex
End Sub

Private Sub ExLVwU_ItemGetInfoTipText(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemGetInfoTipText: listItem=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "ExLVwU_ItemGetInfoTipText: listItem=" & listItem.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If ExLVwU.VirtualMode Then
      infoTipText = "Hello world!"
    Else
      If infoTipText <> "" Then
        infoTipText = infoTipText & vbNewLine & "ID: " & listItem.ID & vbNewLine & "ItemData: " & listItem.ItemData
      Else
        infoTipText = "ID: " & listItem.ID & vbNewLine & "ItemData: " & listItem.ItemData
      End If
    End If
  End If
End Sub

Private Sub ExLVwU_ItemGetOccurrencesCount(ByVal itemIndex As Long, occurrencesCount As Long)
  AddLogEntry "ExLVwU_ItemGetOccurrencesCount: itemIndex=" & itemIndex & ", occurrencesCount=" & occurrencesCount
End Sub

Private Sub ExLVwU_ItemMouseEnter(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemMouseEnter: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_ItemMouseEnter: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_ItemMouseLeave(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemMouseLeave: listItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_ItemMouseLeave: listItem=" & listItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_ItemSelectionChanged(ByVal listItem As ExLVwLibUCtl.IListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemSelectionChanged: listItem=Nothing"
  Else
    AddLogEntry "ExLVwU_ItemSelectionChanged: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwU_ItemSetText(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal itemText As String)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemSetText: listItem=Nothing, itemText=" & itemText
  Else
    AddLogEntry "ExLVwU_ItemSetText: listItem=" & listItem.Index & ", itemText=" & itemText
  End If
End Sub

Private Sub ExLVwU_ItemStateImageChanged(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExLVwLibUCtl.StateImageChangeCausedByConstants)
  Dim li As ExLVwLibUCtl.ListViewItem
  Dim overallStateImageIndex As Long

  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemStateImageChanged: listItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  Else
    AddLogEntry "ExLVwU_ItemStateImageChanged: listItem=" & listItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy

    If bComctl32Version610OrNewer And ExLVwU.ShowHeaderStateImages Then
      overallStateImageIndex = 2
      For Each li In ExLVwU.ListItems
        If li.StateImageIndex = 1 Then
          overallStateImageIndex = 1
          Exit For
        End If
      Next li
      ExLVwU.Columns(0).StateImageIndex = overallStateImageIndex
    End If
  End If
End Sub

Private Sub ExLVwU_ItemStateImageChanging(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExLVwLibUCtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_ItemStateImageChanging: listItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  Else
    AddLogEntry "ExLVwU_ItemStateImageChanging: listItem=" & listItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExLVwU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
  If keyCode = KeyCodeConstants.vbKeyF2 Then
    ExLVwU.CaretItem.StartLabelEditing
  End If
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (ExLVwU.DraggedItems Is Nothing) Then
      ExLVwU.EndDrag True
    End If
  End If
End Sub

Private Sub ExLVwU_KeyPress(keyAscii As Integer)
  AddLogEntry "ExLVwU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExLVwU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ExLVwU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ExLVwU_LostFocus()
  AddLogEntry "ExLVwU_LostFocus"
End Sub

Private Sub ExLVwU_MapGroupWideToTotalItemIndex(ByVal groupIndex As Long, ByVal groupWideItemIndex As Long, totalItemIndex As Long)
  AddLogEntry "ExLVwU_MapGroupWideToTotalItemIndex: groupIndex=" & groupIndex & ", groupWideItemIndex=" & groupWideItemIndex & ", totalItemIndex=" & totalItemIndex
End Sub

Private Sub ExLVwU_MClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_MDblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_MouseDown(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MouseDown: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_MouseEnter(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MouseEnter: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_MouseHover(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MouseHover: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_MouseLeave(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MouseLeave: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_MouseMove(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MouseMove: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

'  AddLogEntry str
End Sub

Private Sub ExLVwU_MouseUp(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_MouseUp: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If button = MouseButtonConstants.vbLeftButton Then
    If Not (ExLVwU.DraggedItems Is Nothing) Then
      ' are we within the client area?
      If ((HitTestConstants.htAbove Or HitTestConstants.htBelow Or HitTestConstants.htToLeft Or HitTestConstants.htToRight) And hitTestDetails) = 0 Then
        ' yes
        ExLVwU.EndDrag False
      Else
        ' no
        ExLVwU.EndDrag True
      End If
    End If
  End If
End Sub

Private Sub ExLVwU_MouseWheel(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, ByVal scrollAxis As ExLVwLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  Dim str As String

  str = "ExLVwU_MouseWheel: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta

  AddLogEntry str
End Sub

Private Sub ExLVwU_OLECompleteDrag(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal performedEffect As ExLVwLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwU_OLEDragDrop(ByVal data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFDIB) Then
    Set draggedPictureU = data.GetData(vbCFDIB)
    ExLVwU.BkImage = draggedPictureU.Handle
  End If
  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ExLVwU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ExLVwU_OLEDragEnter(ByVal data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ExLVwU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ExLVwU_OLEDragLeave(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_OLEDragLeavePotentialTarget()
  AddLogEntry "ExLVwU_OLEDragLeavePotentialTarget"
End Sub

Private Sub ExLVwU_OLEDragMouseMove(ByVal data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExLVwU_OLEGiveFeedback(ByVal effect As ExLVwLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ExLVwU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ExLVwU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As ExLVwLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "ExLVwU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ExLVwU_OLEReceivedNewData(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwU_OLESetData(ByVal data As ExLVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExLVwU_OLEStartDrag(ByVal data As ExLVwLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ExLVwU_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwU_OwnerDrawItem(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal itemState As ExLVwLibUCtl.OwnerDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE)
  Dim str As String

  str = "ExLVwU_OwnerDrawItem: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listItem.Text
  End If
  str = str & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub ExLVwU_RClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_RClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_RDblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_RDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExLVwU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertGroupsU
  InsertItemsU
  If bComctl32Version610OrNewer Then InsertFooterItemsU
End Sub

Private Sub ExLVwU_RemovedColumn(ByVal column As ExLVwLibUCtl.IVirtualListViewColumn)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_RemovedColumn: column=Nothing"
  Else
    AddLogEntry "ExLVwU_RemovedColumn: column=" & column.Caption
  End If
End Sub

Private Sub ExLVwU_RemovedGroup(ByVal Group As ExLVwLibUCtl.IVirtualListViewGroup)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_RemovedGroup: group=Nothing"
  Else
    AddLogEntry "ExLVwU_RemovedGroup: group=" & Group.Text
  End If
End Sub

Private Sub ExLVwU_RemovedItem(ByVal listItem As ExLVwLibUCtl.IVirtualListViewItem)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_RemovedItem: listItem=Nothing"
  Else
    AddLogEntry "ExLVwU_RemovedItem: listItem=" & listItem.Text
  End If
End Sub

Private Sub ExLVwU_RemovingColumn(ByVal column As ExLVwLibUCtl.IListViewColumn, cancelDeletion As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_RemovingColumn: column=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExLVwU_RemovingColumn: column=" & column.Caption & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExLVwU_RemovingGroup(ByVal Group As ExLVwLibUCtl.IListViewGroup, cancelDeletion As Boolean)
  If Group Is Nothing Then
    AddLogEntry "ExLVwU_RemovingGroup: group=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExLVwU_RemovingGroup: group=" & Group.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExLVwU_RemovingItem(ByVal listItem As ExLVwLibUCtl.IListViewItem, cancelDeletion As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_RemovingItem: listItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExLVwU_RemovingItem: listItem=" & listItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExLVwU_RenamedItem(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal previousItemText As String, ByVal newItemText As String)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_RenamedItem: listItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText
  Else
    AddLogEntry "ExLVwU_RenamedItem: listItem=" & listItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText
  End If
End Sub

Private Sub ExLVwU_RenamingItem(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal previousItemText As String, ByVal newItemText As String, cancelRenaming As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_RenamingItem: listItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  Else
    AddLogEntry "ExLVwU_RenamingItem: listItem=" & listItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  End If
End Sub

Private Sub ExLVwU_ResizedControlWindow()
  AddLogEntry "ExLVwU_ResizedControlWindow"
End Sub

Private Sub ExLVwU_ResizingColumn(ByVal column As ExLVwLibUCtl.IListViewColumn, newColumnWidth As Long, Cancel As Boolean)
  If column Is Nothing Then
    AddLogEntry "ExLVwU_ResizingColumn: column=Nothing, newColumnWidth=" & newColumnWidth & ", cancel=" & Cancel
  Else
    AddLogEntry "ExLVwU_ResizingColumn: column=" & column.Caption & ", newColumnWidth=" & newColumnWidth & ", cancel=" & Cancel
  End If
End Sub

Private Sub ExLVwU_SelectedItemRange(ByVal firstItem As ExLVwLibUCtl.IListViewItem, ByVal lastItem As ExLVwLibUCtl.IListViewItem)
  Dim str As String

  str = "ExLVwU_SelectedItemRange: firstItem="
  If firstItem Is Nothing Then
    str = str & "Nothing, lastItem="
  Else
    str = str & firstItem.Text & ", lastItem="
  End If
  If lastItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & lastItem.Text
  End If

  AddLogEntry str
End Sub

Private Sub ExLVwU_SettingItemInfoTipText(ByVal listItem As ExLVwLibUCtl.IListViewItem, infoTipText As String, abortInfoTip As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_SettingItemInfoTipText: listItem=Nothing, infoTipText=" & infoTipText & ", abortInfoTip=" & abortInfoTip
  Else
    AddLogEntry "ExLVwU_SettingItemInfoTipText: listItem=" & listItem.Text & ", infoTipText=" & infoTipText & ", abortInfoTip=" & abortInfoTip
  End If
End Sub

Private Sub ExLVwU_StartingLabelEditing(ByVal listItem As ExLVwLibUCtl.IListViewItem, cancelEditing As Boolean)
  If listItem Is Nothing Then
    AddLogEntry "ExLVwU_StartingLabelEditing: listItem=Nothing, cancelEditing=" & cancelEditing
  Else
    AddLogEntry "ExLVwU_StartingLabelEditing: listItem=" & listItem.Text & ", cancelEditing=" & cancelEditing
  End If
End Sub

Private Sub ExLVwU_SubItemMouseEnter(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwU_SubItemMouseEnter: listSubItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_SubItemMouseEnter: listSubItem=" & listSubItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_SubItemMouseLeave(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If listSubItem Is Nothing Then
    AddLogEntry "ExLVwU_SubItemMouseLeave: listSubItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ExLVwU_SubItemMouseLeave: listSubItem=" & listSubItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ExLVwU_Validate(Cancel As Boolean)
  AddLogEntry "ExLVwU_Validate"
End Sub

Private Sub ExLVwU_XClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_XClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExLVwU_XDblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExLVwU_XDblClick: listItem="
  If listItem Is Nothing Then
    str = str & "Nothing, listSubItem="
  Else
    str = str & listItem.Text & ", listSubItem="
  End If
  If listSubItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & listSubItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
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

  Set wallpaperA = LoadResPicture(102, LoadResConstants.vbResBitmap)
  Set wallpaperU = LoadResPicture(103, LoadResConstants.vbResBitmap)

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
  hImgLst(4) = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst(4), hIcon
      DestroyIcon hIcon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  chkLog.Value = CheckBoxConstants.vbChecked

  Set objActiveCtl = ExLVwU
  UpdateMenu

  InsertColumnsA
  InsertGroupsA
  InsertItemsA
  If bComctl32Version610OrNewer Then InsertFooterItemsA
  InsertColumnsU
  InsertGroupsU
  InsertItemsU
  If bComctl32Version610OrNewer Then InsertFooterItemsU

  If Not bComctl32Version600OrNewer Then
    ' let the listview retrieve the default bitmap's handle
    ExLVwU.Columns(0).BitmapHandle = -1
    hBMPDownArrow = ExLVwU.Columns(0).BitmapHandle
    ExLVwU.Columns(0).BitmapHandle = -2
    hBMPUpArrow = ExLVwU.Columns(0).BitmapHandle
    ExLVwU.Columns(0).BitmapHandle = 0
  End If
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    cx = 0.4 * Me.ScaleWidth
    txtLog.Move Me.ScaleWidth - cx, 0, cx, Me.ScaleHeight - cmdAbout.Height - 10
    cmdAbout.Move txtLog.Left + (cx / 2) - cmdAbout.Width / 2, txtLog.Top + txtLog.Height + 5
    chkLog.Move txtLog.Left, cmdAbout.Top + 5
    ExLVwU.Move 0, 0, txtLog.Left - 5, (Me.ScaleHeight - 5) / 2
    ExLVwA.Move 0, ExLVwU.Top + ExLVwU.Height + 5, ExLVwU.Width, ExLVwU.Height
  End If
End Sub

Private Sub Form_Terminate()
  Dim i As Long

  For i = 1 To 4
    If hImgLst(i) Then ImageList_Destroy hImgLst(i)
  Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If

  If Not bComctl32Version600OrNewer Then
    DeleteObject hBMPDownArrow
    DeleteObject hBMPUpArrow
  End If
End Sub

Private Sub mnuAutoArrange_Click()
  objActiveCtl.AutoArrangeItems = IIf(Not mnuAutoArrange.Checked, AutoArrangeItemsConstants.aaiIntelligent, AutoArrangeItemsConstants.aaiNone)
  mnuAutoArrange.Checked = (objActiveCtl.AutoArrangeItems <> AutoArrangeItemsConstants.aaiNone)
End Sub

Private Sub mnuAutoSizeColumns_Click()
  objActiveCtl.AutoSizeColumns = Not mnuAutoSizeColumns.Checked
  mnuAutoSizeColumns.Checked = objActiveCtl.AutoSizeColumns
End Sub

Private Sub mnuCheckOnSelect_Click()
  objActiveCtl.CheckItemOnSelect = Not mnuCheckOnSelect.Checked
  mnuCheckOnSelect.Checked = objActiveCtl.CheckItemOnSelect
End Sub

Private Sub mnuItemAlignmentLeft_Click()
  objActiveCtl.ItemAlignment = iaLeft
  UpdateMenu
End Sub

Private Sub mnuItemAlignmentTop_Click()
  objActiveCtl.ItemAlignment = iaTop
  UpdateMenu
End Sub

Private Sub mnuJustifyIconColumns_Click()
  objActiveCtl.JustifyIconColumns = Not mnuJustifyIconColumns.Checked
  mnuJustifyIconColumns.Checked = objActiveCtl.JustifyIconColumns
End Sub

Private Sub mnuResizableColumns_Click()
  objActiveCtl.ResizableColumns = Not mnuResizableColumns.Checked
  mnuResizableColumns.Checked = objActiveCtl.ResizableColumns
End Sub

Private Sub mnuShowGroups_Click()
  objActiveCtl.ShowGroups = Not mnuShowGroups.Checked
  mnuShowGroups.Checked = objActiveCtl.ShowGroups
  mnuAutoArrange.Checked = (objActiveCtl.AutoArrangeItems <> AutoArrangeItemsConstants.aaiNone)
End Sub

Private Sub mnuSnapToGrid_Click()
  objActiveCtl.SnapToGrid = Not mnuSnapToGrid.Checked
  mnuSnapToGrid.Checked = objActiveCtl.SnapToGrid
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


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

Private Sub InsertColumnsA()
  ExLVwA.hImageList(ExLVwLibACtl.ImageListConstants.ilHeader) = hImgLst(4)
  With ExLVwA.Columns
    .Add "Column 1", , 160, , , 1, , , True
    .Add "Column 2", , 160, , alCenter, 2, 0
    .Add "Column 3", , 160, , alRight, 3, 0

    .Item(0).IconIndex = 5
    .Item(1).IconIndex = 6
    .Item(2).IconIndex = 7
  End With
End Sub

Private Sub InsertColumnsU()
  ExLVwU.hImageList(ExLVwLibUCtl.ImageListConstants.ilHeader) = hImgLst(4)
  With ExLVwU.Columns
    .Add "Column 1", , 160, , , 1, , , True
    .Add "Column 2", , 160, , alCenter, 2, 0
    .Add "Column 3", , 160, , alRight, 3, 0

    .Item(0).IconIndex = 5
    .Item(1).IconIndex = 6
    .Item(2).IconIndex = 7
  End With
End Sub

Private Sub InsertFooterItemsA()
  If bComctl32Version610OrNewer Then
    ExLVwA.hImageList(ExLVwLibACtl.ImageListConstants.ilFooterItems) = hImgLst(1)
    With ExLVwA.FooterItems
      .Add "Sleep", , 0, 1
      .Add "Rule the world", , 1, 2
    End With
    ExLVwA.ShowFooter
  End If
End Sub

Private Sub InsertFooterItemsU()
  If bComctl32Version610OrNewer Then
    ExLVwU.hImageList(ExLVwLibUCtl.ImageListConstants.ilFooterItems) = hImgLst(1)
    With ExLVwU.FooterItems
      .Add "Sleep", , 0, 1
      .Add "Rule the world", , 1, 2
    End With
    ExLVwU.ShowFooter
  End If
End Sub

Private Sub InsertGroupsA()
  If bComctl32Version610OrNewer Then
    ExLVwA.hImageList(ExLVwLibACtl.ImageListConstants.ilGroups) = hImgLst(1)
  End If
  If bComctl32Version600OrNewer Then
    With ExLVwA.Groups
      With .Add("Group 1", 1, , , , 0, , , "Footer 1", ExLVwLibACtl.AlignmentConstants.alRight, "Subtitle", "Task")
        If bComctl32Version610OrNewer Then
          .DescriptionTextBottom = "Bottom"
          .DescriptionTextTop = "Top"
        End If
      End With
      .Add "Group 2", 2, , , ExLVwLibACtl.AlignmentConstants.alCenter, 1
      .Add "Group 3", 3, , , ExLVwLibACtl.AlignmentConstants.alRight, 2
    End With
  End If
End Sub

Private Sub InsertGroupsU()
  If bComctl32Version610OrNewer Then
    ExLVwU.hImageList(ExLVwLibUCtl.ImageListConstants.ilGroups) = hImgLst(1)
  End If
  If bComctl32Version600OrNewer Then
    With ExLVwU.Groups
      With .Add("Group 1", 1, , , , 0, , , "Footer 1", ExLVwLibUCtl.AlignmentConstants.alRight, "Subtitle", "Task")
        If bComctl32Version610OrNewer Then
          .DescriptionTextBottom = "Bottom"
          .DescriptionTextTop = "Top"
        End If
      End With
      .Add "Group 2", 2, , , ExLVwLibUCtl.AlignmentConstants.alCenter, 1
      .Add "Group 3", 3, , , ExLVwLibUCtl.AlignmentConstants.alRight, 2
    End With
  End If
End Sub

Private Sub InsertItemsA()
  Dim cImages As Long
  Dim iIcon As Long
  Dim itm As ExLVwLibACtl.ListViewItem
  Dim si1a As ExLVwLibACtl.TILEVIEWSUBITEM
  Dim si1b As ExLVwLibACtl.TILEVIEWSUBITEM
  Dim si2a As ExLVwLibACtl.TILEVIEWSUBITEM
  Dim si2b As ExLVwLibACtl.TILEVIEWSUBITEM

  si1a.ColumnIndex = 1
  si1b.ColumnIndex = 1
  si1b.BeginNewColumn = True
  si2a.ColumnIndex = 2
  si2a.BeginNewColumn = True
  si2b.ColumnIndex = 2

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLVwA.hWnd, StrPtr("explorer"), 0
  End If

  Set ExLVwA.BkImage = wallpaperA
  ExLVwA.hImageList(ExLVwLibACtl.ImageListConstants.ilSmall) = hImgLst(1)
  ExLVwA.hImageList(ExLVwLibACtl.ImageListConstants.ilLarge) = hImgLst(2)
  ExLVwA.hImageList(ExLVwLibACtl.ImageListConstants.ilExtraLarge) = hImgLst(3)
  cImages = ImageList_GetImageCount(hImgLst(1))

  If ExLVwA.VirtualMode Then
    ExLVwA.VirtualItemCount = 10
  Else
    With ExLVwA.ListItems
      Set itm = .Add("Item 1", , iIcon, , , 1, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 1, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 1, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 2", , iIcon, , , 2, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 2, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 2, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 3", , iIcon, , , 3, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 3, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 3, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 4", , iIcon, , , 1, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 4, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 4, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 5", , iIcon, , , 2, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 5, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 5, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 6", , iIcon, , , 3, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 6, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 6, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 7", , iIcon, , , 1, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 7, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 7, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 8", , iIcon, , , 2, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 8, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 8, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 9", , iIcon, , , 3, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 9, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 9, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 10", , iIcon, , , 1, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 10, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 10, SubItem 2"
        .IconIndex = iIcon + 2
      End With
    End With
  End If
End Sub

Private Sub InsertItemsU()
  Dim cImages As Long
  Dim iIcon As Long
  Dim itm As ExLVwLibUCtl.ListViewItem
  Dim si1a As ExLVwLibUCtl.TILEVIEWSUBITEM
  Dim si1b As ExLVwLibUCtl.TILEVIEWSUBITEM
  Dim si2a As ExLVwLibUCtl.TILEVIEWSUBITEM
  Dim si2b As ExLVwLibUCtl.TILEVIEWSUBITEM

  si1a.ColumnIndex = 1
  si1b.ColumnIndex = 1
  si1b.BeginNewColumn = True
  si2a.ColumnIndex = 2
  si2a.BeginNewColumn = True
  si2b.ColumnIndex = 2

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLVwU.hWnd, StrPtr("explorer"), 0
  End If

  Set ExLVwU.BkImage = wallpaperU
  ExLVwU.hImageList(ImageListConstants.ilSmall) = hImgLst(1)
  ExLVwU.hImageList(ImageListConstants.ilLarge) = hImgLst(2)
  ExLVwU.hImageList(ImageListConstants.ilExtraLarge) = hImgLst(3)
  cImages = ImageList_GetImageCount(hImgLst(1))

  If ExLVwU.VirtualMode Then
    ExLVwU.VirtualItemCount = 10
  Else
    With ExLVwU.ListItems
      Set itm = .Add("Item 1", , iIcon, , 1, 1, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 1, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 1, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 2", , iIcon, , 2, 2, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 2, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 2, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 3", , iIcon, , 3, 3, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 3, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 3, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 4", , iIcon, , 4, 1, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 4, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 4, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 5", , iIcon, , 5, 2, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 5, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 5, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 6", , iIcon, , 6, 3, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 6, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 6, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 7", , iIcon, , 7, 1, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 7, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 7, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 8", , iIcon, , 8, 2, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 8, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 8, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 9", , iIcon, , 9, 3, Array(si1a, si2a))
      With itm.SubItems(1)
        .Text = "Item 9, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 9, SubItem 2"
        .IconIndex = iIcon + 2
      End With
      iIcon = (iIcon + 1) Mod cImages

      Set itm = .Add("Item 10", , iIcon, , 10, 1, Array(si2b, si1b))
      With itm.SubItems(1)
        .Text = "Item 10, SubItem 1"
        .IconIndex = iIcon + 1
      End With
      With itm.SubItems(2)
        .Text = "Item 10, SubItem 2"
        .IconIndex = iIcon + 2
      End With
    End With
  End If
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
    SendMessageAsLong ExLVwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ExLVwA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
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
  mnuViewTiles.Enabled = bComctl32Version600OrNewer
  mnuViewExtendedTiles.Enabled = bComctl32Version610OrNewer
  Select Case objActiveCtl.ItemAlignment
    Case iaTop
      mnuItemAlignmentTop.Checked = True
      mnuItemAlignmentLeft.Checked = False
    Case iaLeft
      mnuItemAlignmentTop.Checked = False
      mnuItemAlignmentLeft.Checked = True
  End Select
  mnuAutoArrange.Checked = (objActiveCtl.AutoArrangeItems <> AutoArrangeItemsConstants.aaiNone)
  mnuJustifyIconColumns.Checked = objActiveCtl.JustifyIconColumns
  mnuJustifyIconColumns.Enabled = bComctl32Version610OrNewer
  mnuSnapToGrid.Checked = objActiveCtl.SnapToGrid
  mnuSnapToGrid.Enabled = bComctl32Version600OrNewer
  mnuShowGroups.Checked = objActiveCtl.ShowGroups
  mnuShowGroups.Enabled = bComctl32Version600OrNewer
  mnuCheckOnSelect.Checked = objActiveCtl.CheckItemOnSelect
  mnuCheckOnSelect.Enabled = bComctl32Version610OrNewer
  mnuAutoSizeColumns.Checked = objActiveCtl.AutoSizeColumns
  mnuAutoSizeColumns.Enabled = bComctl32Version610OrNewer
  mnuResizableColumns.Checked = objActiveCtl.ResizableColumns
  mnuResizableColumns.Enabled = bComctl32Version610OrNewer
End Sub
