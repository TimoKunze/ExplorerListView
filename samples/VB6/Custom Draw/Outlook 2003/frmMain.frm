VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerListView 1.7 - Outlook 2003 Sample"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   2  'Bildschirmmitte
   Begin ExLVwLibUCtl.ExplorerListView ExLvwU 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _cx             =   6376
      _cy             =   7646
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
      ColumnHeaderVisibility=   0
      DisabledEvents  =   1048317
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
      MultiSelect     =   0   'False
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
      TileViewItemLines=   1
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
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1335
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


  Private Type LOGFONT1
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
'    lfItalic As Byte
'    lfUnderline As Byte
'    lfStrikeOut As Byte
'    lfCharSet As Byte
    lf1 As Long
'    lfOutPrecision As Byte
'    lfClipPrecision As Byte
'    lfQuality As Byte
'    lfPitchAndFamily As Byte
    lf2 As Long
    lfFaceName(1 To 64) As Byte
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private hasFocus As Boolean
  Private hFontBold As Long
  Private hImgLst As Long
  Private ID_Links As Long
  Private ID_MoreLinks As Long
  Private themeableOS As Boolean


  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT1) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal Index As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
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
  ExLvwU.About
End Sub

Private Sub ExLvwU_Click(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If Not (listItem Is Nothing) Then
    If hitTestDetails And HitTestConstants.htItemIcon Then
      Select Case listItem.IconIndex
        Case 0
          listItem.IconIndex = 1
          If listItem.ID = ID_Links Then
            InsertLinks
          ElseIf listItem.ID = ID_MoreLinks Then
            InsertMoreLinks
          End If
        Case 1
          listItem.IconIndex = 0
          If listItem.ID = ID_Links Then
            RemoveLinks
          ElseIf listItem.ID = ID_MoreLinks Then
            RemoveMoreLinks
          End If
      End Select
    End If
  End If
End Sub

Private Sub ExLvwU_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumns
End Sub

Private Sub ExLvwU_CustomDraw(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal drawAllItems As Boolean, textColor As stdole.OLE_COLOR, TextBackColor As stdole.OLE_COLOR, ByVal drawStage As ExLVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExLVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE, furtherProcessing As ExLVwLibUCtl.CustomDrawReturnValuesConstants)
  Const COLOR_HIGHLIGHT = 13
  Const FW_BOLD = 700
  Const WM_GETFONT = &H31
  Dim hFont As Long
  Dim lf As LOGFONT1

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case CustomDrawStageConstants.cdsItemPrePaint
      TextBackColor = ColorConstants.vbWhite
      If listItem.Indent = 0 Then
        If hFontBold = 0 Then
          hFont = SendMessageAsLong(ExLvwU.hWnd, WM_GETFONT, 0, 0)
          GetObjectAPI hFont, LenB(lf), lf
          lf.lfWeight = lf.lfWeight Or FW_BOLD
          hFontBold = CreateFontIndirect(lf)
        End If
        SelectObject hDC, hFontBold

        If Not listItem.Selected Or Not hasFocus Then
          textColor = GetSysColor(COLOR_HIGHLIGHT)
        End If

        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont
      Else
        furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault
      End If
  End Select
End Sub

Private Sub ExLvwU_DblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If Not (listItem Is Nothing) Then
    Select Case listItem.IconIndex
      Case 0
        listItem.IconIndex = 1
        If listItem.ID = ID_Links Then
          InsertLinks
        ElseIf listItem.ID = ID_MoreLinks Then
          InsertMoreLinks
        End If
      Case 1
        listItem.IconIndex = 0
        If listItem.ID = ID_Links Then
          RemoveLinks
        ElseIf listItem.ID = ID_MoreLinks Then
          RemoveMoreLinks
        End If
    End Select
  End If
End Sub

Private Sub ExLvwU_GotFocus()
  hasFocus = True
End Sub

Private Sub ExLvwU_LostFocus()
  hasFocus = False
End Sub

Private Sub ExLvwU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertItems
End Sub

Private Sub ExLvwU_StartingLabelEditing(ByVal listItem As ExLVwLibUCtl.IListViewItem, cancelEditing As Boolean)
  cancelEditing = (listItem.Indent = 0)
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

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 15, 0)
  ImageList_SetImageCount hImgLst, 15
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      Select Case LCase$(iconPath)
        Case "plus.ico"
          ImageList_ReplaceIcon hImgLst, 0, hIcon
        Case "minus.ico"
          ImageList_ReplaceIcon hImgLst, 1, hIcon
        Case "tasks.ico"
          ImageList_ReplaceIcon hImgLst, 2, hIcon
        Case "patterns.ico"
          ImageList_ReplaceIcon hImgLst, 3, hIcon
        Case "deleted items.ico"
          ImageList_ReplaceIcon hImgLst, 4, hIcon
        Case "sent items.ico"
          ImageList_ReplaceIcon hImgLst, 5, hIcon
        Case "journal.ico"
          ImageList_ReplaceIcon hImgLst, 6, hIcon
        Case "junk.ico"
          ImageList_ReplaceIcon hImgLst, 7, hIcon
        Case "calendar.ico"
          ImageList_ReplaceIcon hImgLst, 8, hIcon
        Case "contacts.ico"
          ImageList_ReplaceIcon hImgLst, 9, hIcon
        Case "notes.ico"
          ImageList_ReplaceIcon hImgLst, 10, hIcon
        Case "outgoing.ico"
          ImageList_ReplaceIcon hImgLst, 11, hIcon
        Case "incoming.ico"
          ImageList_ReplaceIcon hImgLst, 12, hIcon
        Case "search.ico"
          ImageList_ReplaceIcon hImgLst, 13, hIcon
        Case Else
          ImageList_AddIcon hImgLst, hIcon
      End Select
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
End Sub

Private Sub Form_Terminate()
  If hFontBold Then DeleteObject hFontBold
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub InsertColumns()
  ExLvwU.Columns.Add "Column 1", , ExLvwU.Width - 10
End Sub

Private Sub InsertItems()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLvwU.hWnd, StrPtr("explorer"), 0
  End If

  ExLvwU.hImageList(ImageListConstants.ilSmall) = hImgLst

  With ExLvwU.ListItems
    ID_Links = .Add("Links", , 1, , 100).ID
    InsertLinks

    ID_MoreLinks = .Add("More Links", , 1, , 100).ID
    InsertMoreLinks
  End With
End Sub

Private Sub InsertLinks()
  Dim baseIndex As Long

  With ExLvwU.ListItems
    baseIndex = .Item(ID_Links, ItemIdentifierTypeConstants.iitID).Index
    .Add "Tasks", baseIndex + 1, 2, 2, ID_Links
    .Add "Patterns", baseIndex + 2, 3, 2, ID_Links
    .Add "Deleted Items", baseIndex + 3, 4, 2, ID_Links
    .Add "Sent Items", baseIndex + 4, 5, 2, ID_Links
  End With
End Sub

Private Sub InsertMoreLinks()
  Dim baseIndex As Long

  With ExLvwU.ListItems
    baseIndex = .Item(ID_MoreLinks, ItemIdentifierTypeConstants.iitID).Index
    .Add "Journal", baseIndex + 1, 6, 2, ID_MoreLinks
    .Add "Junk", baseIndex + 2, 7, 2, ID_MoreLinks
    .Add "Calendar", baseIndex + 3, 8, 2, ID_MoreLinks
    .Add "Contacts", baseIndex + 4, 9, 2, ID_MoreLinks
    .Add "Notes", baseIndex + 5, 10, 2, ID_MoreLinks
    .Add "Outgoing", baseIndex + 6, 11, 2, ID_MoreLinks
    .Add "Incoming", baseIndex + 7, 12, 2, ID_MoreLinks
    .Add "Search", baseIndex + 8, 13, 2, ID_MoreLinks
  End With
End Sub

Private Sub RemoveLinks()
  Dim col As ExLVwLibUCtl.ListViewItems

  Set col = ExLvwU.ListItems
  col.FilterType(FilteredPropertyConstants.fpItemData) = FilterTypeConstants.ftIncluding
  col.Filter(FilteredPropertyConstants.fpItemData) = Array(ID_Links)
  col.RemoveAll
End Sub

Private Sub RemoveMoreLinks()
  Dim col As ExLVwLibUCtl.ListViewItems

  Set col = ExLvwU.ListItems
  col.FilterType(FilteredPropertyConstants.fpItemData) = FilterTypeConstants.ftIncluding
  col.Filter(FilteredPropertyConstants.fpItemData) = Array(ID_MoreLinks)
  col.RemoveAll
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
    SendMessageAsLong ExLvwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
