VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerListView 1.7 - CustomDraw sample"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
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
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkGroups 
      Caption         =   "&Groups"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   975
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLvw 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _cx             =   12303
      _cy             =   7223
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
      DisabledEvents  =   1048191
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
      TextBackColor   =   -2147483643
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
      Left            =   2948
      TabIndex        =   0
      Top             =   4320
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


  Private Const ANSI_CHARSET = 0
  Private Const BALTIC_CHARSET = 186
  Private Const CHINESEBIG5_CHARSET = 136
  Private Const DEFAULT_CHARSET = 1
  Private Const EASTEUROPE_CHARSET = 238
  Private Const GB2312_CHARSET = 134
  Private Const GREEK_CHARSET = 161
  Private Const HANGUL_CHARSET = 129
  Private Const MAC_CHARSET = 77
  Private Const OEM_CHARSET = 255
  Private Const RUSSIAN_CHARSET = 204
  Private Const SHIFTJIS_CHARSET = 128
  Private Const SYMBOL_CHARSET = 2
  Private Const TURKISH_CHARSET = 162


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type ItmData
    ItmType As Byte
    Tag As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private hBMPDownArrow As Long
  Private hBMPUpArrow As Long
  Private hDefaultFont As Long
  Private themeableOS As Boolean
  Private ToolTip As clsToolTip


  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
  Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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


Private Sub chkGroups_Click()
  ExLvw.ShowGroups = (chkGroups.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  ExLvw.About
End Sub

Private Sub ExLvw_ColumnClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  Dim col As ListViewColumn
  Dim curArrow As ExLVwLibUCtl.SortArrowConstants
  Dim newArrow As ExLVwLibUCtl.SortArrowConstants

  For Each col In ExLvw.Columns
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
          ExLvw.SortOrder = soDescending
          ExLvw.SortItems , , , , , col, True
        Case saUp
          ExLvw.SortOrder = soAscending
          ExLvw.SortItems , , , , , col, True
      End Select
    Else
      Select Case newArrow
        Case saNone
          col.BitmapHandle = 0
        Case saDown
          col.BitmapHandle = hBMPDownArrow
          ExLvw.SortOrder = soDescending
          ExLvw.SortItems , , , , , col, True
        Case saUp
          col.BitmapHandle = hBMPUpArrow
          ExLvw.SortOrder = soAscending
          ExLvw.SortItems , , , , , col, True
      End Select
    End If
  Next col
End Sub

Private Sub ExLvw_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumns
End Sub

Private Sub ExLvw_CustomDraw(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal drawAllItems As Boolean, TextColor As stdole.OLE_COLOR, TextBackColor As stdole.OLE_COLOR, ByVal drawStage As ExLVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExLVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE, furtherProcessing As ExLVwLibUCtl.CustomDrawReturnValuesConstants)
  Const OBJ_FONT = 6
  Dim hFont As Long

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case CustomDrawStageConstants.cdsItemPrePaint
      If listItem.Index Mod 2 = 0 Then
        TextBackColor = RGB(240, 240, 240)
      Else
        TextBackColor = ColorConstants.vbWhite
      End If
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont Or CustomDrawReturnValuesConstants.cdrvNotifySubItemDraw

    Case CustomDrawStageConstants.cdsSubItemPrePaint
      If listSubItem.Index > 0 Then
        SelectObject hDC, hDefaultFont
        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont
      Else
        hFont = listItem.ItemData
        If GetObjectType(hFont) = OBJ_FONT Then SelectObject hDC, hFont
        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont
      End If
  End Select
End Sub

Private Sub ExLvw_DestroyedControlWindow(ByVal hWnd As Long)
  ToolTip.Detach
End Sub

Private Sub ExLvw_FreeItemData(ByVal listItem As ExLVwLibUCtl.IListViewItem)
  Const OBJ_FONT = 6
  Dim h As Long

  h = listItem.ItemData
  If GetObjectType(h) = OBJ_FONT Then DeleteObject h
End Sub

Private Sub ExLvw_ItemGetInfoTipText(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Const OBJ_FONT = 6

  If GetObjectType(listItem.ItemData) = OBJ_FONT Then
    infoTipText = listItem.Text
  End If
End Sub

Private Sub ExLvw_RecreatedControlWindow(ByVal hWnd As Long)
  InsertItems
End Sub

Private Sub Form_Initialize()
  Dim DLLVerData As DLLVERSIONINFO
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  Set ToolTip = New clsToolTip

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  InsertColumns

  If Not bComctl32Version600OrNewer Then
    chkGroups.Enabled = False
    ' let the listview retrieve the default bitmap's handle
    ExLvw.Columns(0).BitmapHandle = -1
    hBMPDownArrow = ExLvw.Columns(0).BitmapHandle
    ExLvw.Columns(0).BitmapHandle = -2
    hBMPUpArrow = ExLvw.Columns(0).BitmapHandle
    ExLvw.Columns(0).BitmapHandle = 0
  End If

  InsertItems
End Sub

Private Sub Form_Unload(cancel As Integer)
  Const OBJ_FONT = 6
  Dim h As Long
  Dim listItem As ListViewItem

  ' The FreeItemData won't get fired on program termination (actually it gets fired, but we won't
  ' receive it anymore). So ensure the fonts get freed.
  For Each listItem In ExLvw.ListItems
    h = listItem.ItemData
    If GetObjectType(h) = OBJ_FONT Then
      DeleteObject h
    End If
  Next listItem
  ToolTip.Detach
  Set ToolTip = Nothing

  If Not bComctl32Version600OrNewer Then
    DeleteObject hBMPDownArrow
    DeleteObject hBMPUpArrow
  End If

  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub InsertColumns()
  With ExLvw.Columns
    .Add("Font name", , 270).ImageOnRight = True
    .Add("Charset", , 150).ImageOnRight = True
  End With
End Sub

Private Sub InsertItems()
  Const WM_GETFONT = &H31
  Dim hDC As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLvw.hWnd, StrPtr("explorer"), 0
  End If

  ToolTip.Attach ExLvw.hWndToolTip
  ToolTip.BalloonStyle = True
  ToolTip.TitleText = "Font name:"
  ToolTip.TitleIcon = ToolTipTitleIconConstants.tttiInfo

  If bComctl32Version600OrNewer Then
    With ExLvw.Groups
      .Add "ANSI_CHARSET", ANSI_CHARSET, , , , , True
      .Add "BALTIC_CHARSET", BALTIC_CHARSET, , , , , True
      .Add "CHINESEBIG5_CHARSET", CHINESEBIG5_CHARSET, , , , , True
      .Add "DEFAULT_CHARSET", DEFAULT_CHARSET, , , , , True
      .Add "EASTEUROPE_CHARSET", EASTEUROPE_CHARSET, , , , , True
      .Add "GB2312_CHARSET", GB2312_CHARSET, , , , , True
      .Add "GREEK_CHARSET", GREEK_CHARSET, , , , , True
      .Add "HANGUL_CHARSET", HANGUL_CHARSET, , , , , True
      .Add "MAC_CHARSET", MAC_CHARSET, , , , , True
      .Add "OEM_CHARSET", OEM_CHARSET, , , , , True
      .Add "RUSSIAN_CHARSET", RUSSIAN_CHARSET, , , , , True
      .Add "SHIFTJIS_CHARSET", SHIFTJIS_CHARSET, , , , , True
      .Add "SYMBOL_CHARSET", SYMBOL_CHARSET, , , , , True
      .Add "TURKISH_CHARSET", TURKISH_CHARSET, , , , , True
    End With
  End If

  With ExLvw.ListItems
    hDefaultFont = SendMessage(ExLvw.hWnd, WM_GETFONT, 0, 0)
    GetObjectAPI hDefaultFont, LenB(lfListView), lfListView
    hDC = GetDC(0)

    EnumFonts hDC, ExLvw, ANSI_CHARSET, "ANSI_CHARSET"
    EnumFonts hDC, ExLvw, BALTIC_CHARSET, "BALTIC_CHARSET"
    EnumFonts hDC, ExLvw, CHINESEBIG5_CHARSET, "CHINESEBIG5_CHARSET"
    EnumFonts hDC, ExLvw, DEFAULT_CHARSET, "DEFAULT_CHARSET"
    EnumFonts hDC, ExLvw, EASTEUROPE_CHARSET, "EASTEUROPE_CHARSET"
    EnumFonts hDC, ExLvw, GB2312_CHARSET, "GB2312_CHARSET"
    EnumFonts hDC, ExLvw, GREEK_CHARSET, "GREEK_CHARSET"
    EnumFonts hDC, ExLvw, HANGUL_CHARSET, "HANGUL_CHARSET"
    EnumFonts hDC, ExLvw, MAC_CHARSET, "MAC_CHARSET"
    EnumFonts hDC, ExLvw, OEM_CHARSET, "OEM_CHARSET"
    EnumFonts hDC, ExLvw, RUSSIAN_CHARSET, "RUSSIAN_CHARSET"
    EnumFonts hDC, ExLvw, SHIFTJIS_CHARSET, "SHIFTJIS_CHARSET"
    EnumFonts hDC, ExLvw, SYMBOL_CHARSET, "SYMBOL_CHARSET"
    EnumFonts hDC, ExLvw, TURKISH_CHARSET, "TURKISH_CHARSET"

    ReleaseDC 0, hDC
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
    SendMessageAsLong ExLvw.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
