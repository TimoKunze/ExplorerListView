VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerListView 1.7 - SubItemControls Sample"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
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
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   669
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkInPlaceEditing 
      Caption         =   "&In-Place Editing"
      Height          =   255
      Left            =   8490
      TabIndex        =   2
      Top             =   3960
      Value           =   1  'Aktiviert
      Width           =   1575
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
      Left            =   8610
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwU 
      Align           =   3  'Links ausrichten
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
      _cx             =   14790
      _cy             =   7567
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   0
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
      OLEDragImageStyle=   1
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
      View            =   3
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmMain.frx":0000
      FooterIntroText =   "frmMain.frx":0072
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

  Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private bComctl32Version610OrNewer As Boolean
  Private hImgLst(1 To 3) As Long
  Private IID_IPropertyDescription As GUID
  Private themeableOS As Boolean


  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As GUID) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
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
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function PropVariantToString Lib "propsys.dll" (ByVal propvar As Long, ByVal psz As Long, ByVal cch As Long) As Long
  Private Declare Function PropVariantToVariant Lib "propsys.dll" (ByVal pPropVar As Long, ByVal pVar As Long) As Long
  Private Declare Function PSGetPropertyDescriptionByName Lib "propsys.dll" (ByVal pszCanonicalName As Long, riid As GUID, ppv As Long) As Long
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

Private Sub ExLVwU_ConfigureSubItemControl(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal controlKind As ExLVwLibUCtl.SubItemControlKindConstants, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants, ByVal subItemControl As ExLVwLibUCtl.SubItemControlConstants, themeAppName As String, themeIDList As String, hFont As Long, textColor As stdole.OLE_COLOR, pPropertyDescription As Long, ByVal pPropertyValue As Long)
  Const VT_BOOL As Long = 11
  Const VT_DATE As Long = 7
  Const VT_I4 As Long = 3
  Dim dDate As Date
  Dim lVal As Long

  ' NOTE: With some more effort, you could provide custom implementations of the IPropertyDescription interface
  Select Case subItemControl
    Case sicDropList
      PSGetPropertyDescriptionByName StrPtr("System.Photo.MeteringMode"), IID_IPropertyDescription, pPropertyDescription
      CopyMemory pPropertyValue, VarPtr(VT_I4), 2
      lVal = Val(listSubItem.Text)
      CopyMemory pPropertyValue + 8, VarPtr(lVal), LenB(lVal)
  End Select
End Sub

Private Sub ExLVwU_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumns
End Sub

Private Sub ExLVwU_EndSubItemEdit(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants, ByVal pPropertyKey As Long, ByVal pPropertyValue As Long, Cancel As Boolean)
  Dim propValue As String

  propValue = String$(1024, 0)
  PropVariantToString pPropertyValue, StrPtr(propValue), 1024
  propValue = Left$(propValue, lstrlen(StrPtr(propValue)))
  listSubItem.Text = propValue
End Sub

Private Sub ExLVwU_GetSubItemControl(ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal controlKind As ExLVwLibUCtl.SubItemControlKindConstants, ByVal editMode As ExLVwLibUCtl.SubItemEditModeConstants, subItemControl As ExLVwLibUCtl.SubItemControlConstants)
  If Not (listSubItem Is Nothing) Then
    If listSubItem.Index = 1 Then
      If (controlKind <> sickInPlaceEditing) Or (chkInPlaceEditing.Value = vbChecked) Then
        subItemControl = listSubItem.ParentItem.ItemData
      End If
    End If
  End If
End Sub

Private Sub ExLVwU_InvokeVerbFromSubItemControl(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal verb As String)
  MsgBox "TODO for item '" & listItem.Text & "': " & verb
End Sub

Private Sub ExLVwU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertItems
  UpdateMenu
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Const strIID_IPropertyDescription As String = "{6f79d558-3e96-4549-a1d1-7d75d2288814}"
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim hMod As Long
  Dim i As Long
  Dim iconsDir As String
  Dim iconPath As String
  Dim iconSize As Long

  InitCommonControls
  CLSIDFromString StrPtr(strIID_IPropertyDescription), IID_IPropertyDescription

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
    If .dwMajor >= 6 Then
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

  InsertColumns
  InsertItems
  UpdateMenu
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

Private Sub mnuViewDetails_Click()
  ExLVwU.View = vDetails
  ExLVwU.FullRowSelect = frsExtendedMode
  UpdateMenu
End Sub

Private Sub mnuViewExtendedTiles_Click()
  ExLVwU.View = vExtendedTiles
  ExLVwU.FullRowSelect = frsDisabled
  UpdateMenu
End Sub

Private Sub mnuViewIcons_Click()
  ExLVwU.View = vIcons
  UpdateMenu
End Sub

Private Sub mnuViewList_Click()
  ExLVwU.View = vList
  UpdateMenu
End Sub

Private Sub mnuViewSmallIcons_Click()
  ExLVwU.View = vSmallIcons
  UpdateMenu
End Sub

Private Sub mnuViewTiles_Click()
  ExLVwU.View = vTiles
  ExLVwU.FullRowSelect = frsDisabled
  UpdateMenu
End Sub


Private Sub InsertColumns()
  With ExLVwU.Columns
    .Add "Column 1", , 160, , , 1
    .Add "Column 2", , 200, , , 2
    .Add "Column 3", , 160, , , 3
  End With
End Sub

Private Sub InsertItems()
  Dim cImages As Long
  Dim iIcon As Long
  Dim itm As ListViewItem
  Dim si1a As TILEVIEWSUBITEM
  Dim si1b As TILEVIEWSUBITEM
  Dim si2a As TILEVIEWSUBITEM
  Dim si2b As TILEVIEWSUBITEM

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

  ExLVwU.hImageList(ImageListConstants.ilSmall) = hImgLst(1)
  ExLVwU.hImageList(ImageListConstants.ilLarge) = hImgLst(2)
  ExLVwU.hImageList(ImageListConstants.ilExtraLarge) = hImgLst(3)
  cImages = ImageList_GetImageCount(hImgLst(1))

  With ExLVwU.ListItems
    Set itm = .Add("No Sub-Item Control", , iIcon, , , , Array(si1a, si2a))
    itm.SubItems(1).Text = "Hello World"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Multi-Line Text", , iIcon, , sicMultiLineText, , Array(si2b, si1b))
    itm.SubItems(1).Text = "This text" & vbNewLine & "consists of" & vbNewLine & "3 lines"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Percent Bar", , iIcon, , sicPercentBar, , Array(si1a, si2a))
    itm.SubItems(1).Text = "46"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Rating", , iIcon, , sicRating, , Array(si1a, si2a))
    itm.SubItems(1).Text = "3"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Text", , iIcon, , sicText, , Array(si2b, si1b))
    itm.SubItems(1).Text = "Some Text"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Boolean Check Mark", , iIcon, , sicBooleanCheckMark, , Array(si2b, si1b))
    itm.SubItems(1).Text = "-1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Calendar", , iIcon, , sicCalendar, , Array(si1a, si2a))
    itm.SubItems(1).Text = Date
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Boolean Drop-Down List", , iIcon, , sicCheckboxDropList, , Array(si2b, si1b))
    itm.SubItems(1).Text = "-1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Drop-Down List", , iIcon, , sicDropList, , Array(si2b, si1b))
    itm.SubItems(1).Text = "2"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Hyperlink", , iIcon, , sicHyperlink, , Array(si2b, si1b))
    itm.SubItems(1).Text = "<a id=""Open https://www.timosoft-software.de"">https://www.timosoft-software.de</a>"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Normal Sub-Item"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
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
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong ExLVwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Sub UpdateMenu()
  Select Case ExLVwU.View
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
End Sub
