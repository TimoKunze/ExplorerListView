VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerListView 1.7 - State Images Sample"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8115
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
   ScaleWidth      =   541
   StartUpPosition =   2  'Bildschirmmitte
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
      Left            =   6600
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwU 
      Align           =   3  'Links ausrichten
      Height          =   6555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _cx             =   11245
      _cy             =   11562
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
      CallBackMask    =   61440
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
      FullRowSelect   =   0
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
      View            =   2
      VirtualMode     =   -1  'True
      EmptyMarkupText =   "frmMain.frx":0000
      FooterIntroText =   "frmMain.frx":0072
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Details"
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


  Private Const cItems = 2000


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type


  Private arrStateImageIndices(cItems - 1) As Long
  Private bComctl32Version600OrNewer As Boolean
  Private cImages As Long
  Private hImgLst As Long
  Private themeableOS As Boolean


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
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
  ExLVwU.About
End Sub

Private Sub ExLVwU_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumns
End Sub

Private Sub ExLVwU_ItemGetDisplayInfo(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal requestedInfo As ExLVwLibUCtl.RequestedInfoConstants, IconIndex As Long, Indent As Long, groupID As Long, TileViewColumns() As ExLVwLibUCtl.TILEVIEWSUBITEM, ByVal maxItemTextLength As Long, itemText As String, OverlayIndex As Long, StateImageIndex As Long, itemStates As ExLVwLibUCtl.ItemStateConstants, dontAskAgain As Boolean)
  Dim c(1 To 2) As ExLVwLibUCtl.TILEVIEWSUBITEM

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
  If requestedInfo And RequestedInfoConstants.riStateImageIndex Then
    StateImageIndex = arrStateImageIndices(listItem.Index) + 1
  End If
End Sub

Private Sub ExLVwU_ItemStateImageChanged(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExLVwLibUCtl.StateImageChangeCausedByConstants)
  arrStateImageIndices(listItem.Index) = newStateImageIndex - 1
  listItem.Update
End Sub

Private Sub ExLVwU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertItems
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
  End With

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & IIf(bComctl32Version600OrNewer, "16x16x32bpp\", "16x16x8bpp\")
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  InsertColumns
  InsertItems
  UpdateMenu
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub mnuViewDetails_Click()
  ExLVwU.View = vDetails
  ExLVwU.FullRowSelect = frsExtendedMode
  UpdateMenu
End Sub

Private Sub mnuViewList_Click()
  ExLVwU.View = vList
  ExLVwU.FullRowSelect = frsDisabled
  UpdateMenu
End Sub


Private Sub InsertColumns()
  With ExLVwU.Columns
    .Add "Column 1", , 160, , , 1
    .Add "Column 2", , 160, , alCenter, 2
  End With
End Sub

Private Sub InsertItems()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLVwU.hWnd, StrPtr("explorer"), 0
  End If

  ExLVwU.hImageList(ImageListConstants.ilSmall) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)
 
  ExLVwU.VirtualItemCount = cItems
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong ExLVwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Sub UpdateMenu()
  Select Case ExLVwU.View
    Case vDetails
      mnuViewDetails.Checked = True
      mnuViewList.Checked = False
    Case vList
      mnuViewDetails.Checked = False
      mnuViewList.Checked = True
  End Select
End Sub
