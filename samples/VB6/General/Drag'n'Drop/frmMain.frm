VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerListView 1.7 - Drag'n'Drop Sample"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
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
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkOLEDragDrop 
      Caption         =   "&OLE Drag'n'Drop"
      Height          =   255
      Left            =   6570
      TabIndex        =   2
      Top             =   3600
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
      Left            =   6690
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLVwU 
      Align           =   3  'Links ausrichten
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _cx             =   11404
      _cy             =   7011
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
      DisabledEvents  =   1048054
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
      RegisterForOLEDragDrop=   -1  'True
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
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoArrange 
         Caption         =   "&Auto-arrange"
      End
      Begin VB.Menu mnuSnapToGrid 
         Caption         =   "Snap To G&rid"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAeroDragImages 
         Caption         =   "Aer&o OLE Drag Images"
         Checked         =   -1  'True
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
   Begin VB.Menu mnuPopup 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy item"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move item"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "C&ancel"
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


  Private Enum ChosenMenuItemConstants
    ciCopy
    ciMove
    ciCancel
  End Enum


  Private Const CF_DIBV5 = 17
  Private Const CF_OEMTEXT = 7
  Private Const CF_UNICODETEXT = 13

  Private Const strCLSID_RecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}"


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

  Private Type POINTAPI
    x As Long
    y As Long
  End Type

  Private Type Size
    cx As Long
    cy As Long
  End Type


  Private CFSTR_HTML As Long
  Private CFSTR_HTML2 As Long
  Private CFSTR_LOGICALPERFORMEDDROPEFFECT As Long
  Private CFSTR_MOZILLAHTMLCONTEXT As Long
  Private CFSTR_PERFORMEDDROPEFFECT As Long
  Private CFSTR_SHELLIDLISTOFFSET As Long
  Private CFSTR_TARGETCLSID As Long
  Private CFSTR_TEXTHTML As Long

  Private bComctl32Version600OrNewer As Boolean
  Private bComctl32Version610OrNewer As Boolean
  Private bRightDrag As Boolean
  Private chosenMenuItem As ChosenMenuItemConstants
  Private CLSID_RecycleBin As GUID
  Private draggedPicture As StdPicture
  Private freeFloat As Boolean
  Private hImgLst(1 To 3) As Long
  Private lastDropTarget As Long
  Private ptStartDrag As POINTAPI
  Private szHotSpotOffset As Size
  Private themeableOS As Boolean
  Private ToolTip As clsToolTip


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
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long
  Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
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

Private Sub ExLVwU_AbortedDrag()
  Set ExLVwU.DropHilitedItem = Nothing
  lastDropTarget = -1
  If bComctl32Version600OrNewer And Not freeFloat Then
    ExLVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  End If
End Sub

Private Sub ExLVwU_ColumnBeginDrag(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, doAutomaticDragDrop As Boolean)
  bRightDrag = False
  lastDropTarget = -1

  If Not (column Is Nothing) Then
    If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
      doAutomaticDragDrop = False
      ExLVwU.HeaderOLEDrag , , , column
    Else
      doAutomaticDragDrop = False
      ExLVwU.HeaderBeginDrag column, -1, szHotSpotOffset.cx, szHotSpotOffset.cy
    End If
  End If
End Sub

Private Sub ExLVwU_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumns
End Sub

Private Sub ExLVwU_DragMouseMove(dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim insertMarkPosition As InsertMarkPositionConstants
  Dim itm As ListViewItem
  Dim newDropTarget As Long

  If dropTarget Is Nothing Then
    newDropTarget = -1
  Else
    newDropTarget = dropTarget.Index
  End If
  lastDropTarget = newDropTarget

  If bComctl32Version600OrNewer And Not freeFloat Then
    ExLVwU.GetClosestInsertMarkPosition Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), insertMarkPosition, itm
    ExLVwU.SetInsertMarkPosition insertMarkPosition, itm
  End If
End Sub

Private Sub ExLVwU_Drop(ByVal dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
'  Dim alignLeft As Boolean
  Dim dx As Long
  Dim dy As Long
  Dim itm As ListViewItem
  Dim xPos As Long
  Dim yPos As Long

  x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  If bRightDrag Then
    Me.PopupMenu mnuPopup, , x, y, mnuMove
    Select Case chosenMenuItem
      Case ChosenMenuItemConstants.ciCancel
        ExLVwU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciCopy
        MsgBox "ToDo: Copy the item.", VbMsgBoxStyle.vbInformation, "Command not implemented"
        ExLVwU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciMove
        ' fall through
    End Select
  End If

'  alignLeft = (ExLVwU.ItemAlignment = ItemAlignmentConstants.iaLeft)
  dx = x - ptStartDrag.x
  dy = y - ptStartDrag.y

  For Each itm In ExLVwU.DraggedItems
    If freeFloat Then
      itm.GetPosition xPos, yPos
      itm.SetPosition xPos + dx, yPos + dy
    Else
      ' TODO
'      itm.GetPosition xPos, yPos
'      If alignLeft Then
'        itm.SetPosition xPos + dx, y
'      Else
'        itm.SetPosition x, yPos + dy
'      End If
    End If
  Next itm

  lastDropTarget = -1
  If bComctl32Version600OrNewer And Not freeFloat Then
    ExLVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  End If
End Sub

Private Sub ExLVwU_HeaderAbortedDrag()
  ExLVwU.SetHeaderInsertMarkPosition -1
  lastDropTarget = -1
  ExLVwU.HeaderRefresh
End Sub

Private Sub ExLVwU_HeaderDragMouseMove(dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim newDropTarget As Long

  If dropTarget Is Nothing Then
    newDropTarget = -1
    ExLVwU.SetHeaderInsertMarkPosition -1
  Else
    newDropTarget = dropTarget.Index
    ExLVwU.SetHeaderInsertMarkPosition , Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  End If
  lastDropTarget = newDropTarget
  ExLVwU.HeaderRefresh
End Sub

Private Sub ExLVwU_HeaderDrop(ByVal dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  Dim p As Long

  If bRightDrag Then
    x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    Me.PopupMenu mnuPopup, , x, y, mnuMove
    Select Case chosenMenuItem
      Case ChosenMenuItemConstants.ciCancel
        ExLVwU_HeaderAbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciCopy
        MsgBox "ToDo: Copy the column.", VbMsgBoxStyle.vbInformation, "Command not implemented"
        ExLVwU_HeaderAbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciMove
        ' fall through
    End Select
  End If

  ' move the column

  p = ExLVwU.SetHeaderInsertMarkPosition(, Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels))
  With ExLVwU.DraggedColumn
    If .Position < p Then
      .Position = p - 1
    Else
      .Position = p
    End If
  End With

  ExLVwU.SetHeaderInsertMarkPosition -1
  lastDropTarget = -1
  ExLVwU.HeaderRefresh
End Sub

Private Sub ExLVwU_HeaderMouseUp(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  If button = MouseButtonConstants.vbLeftButton Then
    If Not (ExLVwU.DraggedColumn Is Nothing) Then
      ' are we within the (extended) client area?
      If lastDropTarget <> -1 Then
        ' yes
        ExLVwU.HeaderEndDrag False
      Else
        ' no
        ExLVwU.HeaderEndDrag True
      End If
    End If
  End If
End Sub

Private Sub ExLVwU_HeaderOLECompleteDrag(ByVal Data As ExLVwLibUCtl.IOLEDataObject, ByVal performedEffect As ExLVwLibUCtl.OLEDropEffectConstants)
  Dim bytArray() As Byte
  Dim CLSIDOfTarget As GUID

  If Data.GetFormat(CFSTR_TARGETCLSID) Then
    bytArray = Data.GetData(CFSTR_TARGETCLSID)
    CopyMemory VarPtr(CLSIDOfTarget), VarPtr(bytArray(LBound(bytArray))), LenB(CLSIDOfTarget)
    If IsEqualGUID(CLSIDOfTarget, CLSID_RecycleBin) Then
      ' remove the dragged column
      ExLVwU.Columns.Remove ExLVwU.DraggedColumn.Index
    End If
  End If
End Sub

Private Sub ExLVwU_HeaderOLEDragDrop(ByVal Data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, ByVal dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  Dim p As Long
  Dim str As String

  If Data.GetFormat(vbCFDIB) Then
    Set draggedPicture = Data.GetData(vbCFDIB)
    ExLVwU.BkImage = draggedPicture.Handle
  Else
    If Data.GetFormat(CF_UNICODETEXT) Then
      str = Data.GetData(CF_UNICODETEXT)
    ElseIf Data.GetFormat(vbCFText) Then
      str = Data.GetData(vbCFText)
    ElseIf Data.GetFormat(CF_OEMTEXT) Then
      str = Data.GetData(CF_OEMTEXT)
    End If
    str = Replace$(str, vbNewLine, "")

    If str <> "" Then
      ' insert a new column
      p = ExLVwU.SetHeaderInsertMarkPosition(, Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels))
      ExLVwU.Columns.Add str, p
    End If
  End If

  ExLVwU.SetHeaderInsertMarkPosition -1
  lastDropTarget = -1

  ' Be careful with odeMove!! Some sources delete the data themselves if the Move effect was
  ' chosen. So you may lose data!
  'If shift And vbShiftMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeLink
End Sub

Private Sub ExLVwU_HeaderOLEDragEnter(ByVal Data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim newDropTarget As Long

  If Not Data.GetFormat(vbCFDIB) Then
    If dropTarget Is Nothing Then
      newDropTarget = -1
      ExLVwU.SetHeaderInsertMarkPosition -1
    Else
      newDropTarget = dropTarget.Index
      ExLVwU.SetHeaderInsertMarkPosition , Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    End If
    lastDropTarget = newDropTarget
  End If

  effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
'  If shift And vbShiftMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeMove
'  If shift And vbCtrlMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
'  If shift And vbAltMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeLink
  Data.SetDropDescription , "Insert as new column here", DropDescriptionIconConstants.ddiCopy
End Sub

Private Sub ExLVwU_HeaderOLEDragLeave(ByVal Data As ExLVwLibUCtl.IOLEDataObject, ByVal dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
  ExLVwU.SetHeaderInsertMarkPosition -1
  lastDropTarget = -1
  Data.SetDropDescription , , DropDescriptionIconConstants.ddiUseDefault
End Sub

Private Sub ExLVwU_HeaderOLEDragMouseMove(ByVal Data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal xListView As Single, ByVal yListView As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim newDropTarget As Long

  If Not Data.GetFormat(vbCFDIB) Then
    If dropTarget Is Nothing Then
      newDropTarget = -1
      ExLVwU.SetHeaderInsertMarkPosition -1
    Else
      newDropTarget = dropTarget.Index
      ExLVwU.SetHeaderInsertMarkPosition , Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    End If
    lastDropTarget = newDropTarget
  End If

  effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
'  If shift And vbShiftMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeMove
'  If shift And vbCtrlMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
'  If shift And vbAltMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeLink
  Data.SetDropDescription , "Insert as new column here", DropDescriptionIconConstants.ddiCopy
End Sub

Private Sub ExLVwU_HeaderOLESetData(ByVal Data As ExLVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  Select Case formatID
    Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
      Data.SetData formatID, ExLVwU.DraggedColumn.Caption

    Case vbCFFiles
      ' WARNING! Dragging files to objects like the recycle bin may result in deleted files and
      ' lost data!
'      ReDim files(1 To 2) As String
'      files(1) = "C:\Test1.txt"
'      files(2) = "C:\Test2.txt"
'      Data.SetData formatID, files

    Case vbCFBitmap, vbCFDIB, CF_DIBV5
      Data.SetData formatID, draggedPicture

    Case vbCFPalette
      If Not (draggedPicture Is Nothing) Then
        Data.SetData formatID, draggedPicture.hPal
      End If
  End Select
End Sub

Private Sub ExLVwU_HeaderOLEStartDrag(ByVal Data As ExLVwLibUCtl.IOLEDataObject)
  Data.SetData vbCFText
  Data.SetData CF_OEMTEXT
  Data.SetData CF_UNICODETEXT
  Data.SetData vbCFFiles
  If Not (draggedPicture Is Nothing) Then
    Data.SetData vbCFPalette
    Data.SetData vbCFBitmap
    Data.SetData vbCFDIB
    Data.SetData CF_DIBV5
  End If
End Sub

Private Sub ExLVwU_ItemBeginDrag(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim col As ListViewItems

  bRightDrag = False
  lastDropTarget = -1
  Set col = ExLVwU.ListItems
  col.FilterType(FilteredPropertyConstants.fpSelected) = FilterTypeConstants.ftIncluding
  col.Filter(FilteredPropertyConstants.fpSelected) = Array(True)

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
    ExLVwU.OLEDrag , , , ExLVwU.CreateItemContainer(col)
  Else
    ExLVwU.BeginDrag ExLVwU.CreateItemContainer(col), -1, szHotSpotOffset.cx, szHotSpotOffset.cy
  End If
End Sub

Private Sub ExLVwU_ItemBeginRDrag(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim col As ListViewItems

  bRightDrag = True
  lastDropTarget = -1
  Set col = ExLVwU.ListItems
  col.FilterType(FilteredPropertyConstants.fpSelected) = FilterTypeConstants.ftIncluding
  col.Filter(FilteredPropertyConstants.fpSelected) = Array(True)

  ptStartDrag.x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  ptStartDrag.y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
    ExLVwU.OLEDrag , , , ExLVwU.CreateItemContainer(col)
  Else
    ExLVwU.BeginDrag ExLVwU.CreateItemContainer(col), -1, szHotSpotOffset.cx, szHotSpotOffset.cy
  End If
End Sub

Private Sub ExLVwU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (ExLVwU.DraggedItems Is Nothing) Then ExLVwU.EndDrag True
    If Not (ExLVwU.DraggedColumn Is Nothing) Then ExLVwU.HeaderEndDrag True
  End If
End Sub

Private Sub ExLVwU_MouseUp(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  If chkOLEDragDrop.Value = CheckBoxConstants.vbUnchecked Then
    If ((button = MouseButtonConstants.vbLeftButton) And (Not bRightDrag)) Or ((button = MouseButtonConstants.vbRightButton) And bRightDrag) Then
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
  End If
End Sub

Private Sub ExLVwU_OLECompleteDrag(ByVal Data As ExLVwLibUCtl.IOLEDataObject, ByVal performedEffect As ExLVwLibUCtl.OLEDropEffectConstants)
  Dim bytArray() As Byte
  Dim CLSIDOfTarget As GUID

  If Data.GetFormat(CFSTR_TARGETCLSID) Then
    bytArray = Data.GetData(CFSTR_TARGETCLSID)
    CopyMemory VarPtr(CLSIDOfTarget), VarPtr(bytArray(LBound(bytArray))), LenB(CLSIDOfTarget)
    If IsEqualGUID(CLSIDOfTarget, CLSID_RecycleBin) Then
      ' remove the dragged items
      ExLVwU.DraggedItems.RemoveAll True
    End If
  End If
End Sub

Private Sub ExLVwU_OLEDragDrop(ByVal Data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  Dim bytArray() As Byte
  Dim files() As String
  Dim hasPositions As Boolean
  Dim i As Long
  Dim itm As ListViewItem
  Dim origin As POINTAPI
  Dim p As Long
  Dim positions() As POINTAPI
  Dim pt As POINTAPI
  Dim spacingLarge As Size
  Dim spacingSmall As Size
  Dim str As String

  x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  If Data.GetFormat(vbCFDIB) Then
    Set draggedPicture = Data.GetData(vbCFDIB)
    ExLVwU.BkImage = draggedPicture.Handle
  Else
    If Data.GetFormat(CF_UNICODETEXT) Then
      str = Data.GetData(CF_UNICODETEXT)
    ElseIf Data.GetFormat(vbCFText) Then
      str = Data.GetData(vbCFText)
    ElseIf Data.GetFormat(CF_OEMTEXT) Then
      str = Data.GetData(CF_OEMTEXT)
    End If
    str = Replace$(str, vbNewLine, "")

    If str <> "" Then
      ' insert a new item
      Set itm = ExLVwU.ListItems.Add(str, , 1)
      itm.SetPosition x, y
    ElseIf Data.GetFormat(vbCFFiles) Then
      ' insert a new item for each file/folder
      On Error GoTo NoFiles
      files = Data.GetData(vbCFFiles)
      If Data.GetFormat(CFSTR_SHELLIDLISTOFFSET) Then
        hasPositions = True
        ReDim positions(LBound(files) To UBound(files)) As POINTAPI
        bytArray = Data.GetData(CFSTR_SHELLIDLISTOFFSET)

        ' the first POINTAPI structure specifies the position (in client coordinates of the source
        ' window) of the mouse cursor within the drag image
        CopyMemory VarPtr(pt), VarPtr(bytArray(1)), LenB(pt)
        ' the remainder of the POINTAPI structures specify the locations of the individual files relative to
        ' the position specified by 'pt'
        CopyMemory VarPtr(positions(LBound(positions))), VarPtr(bytArray(1 + LenB(pt))), (UBound(files) - LBound(files) + 1) * LenB(pt)

        If ExLVwU.View = ViewConstants.vIcons Then
          spacingLarge.cx = 1
          spacingLarge.cy = 1
          spacingSmall.cx = 1
          spacingSmall.cy = 1
        ElseIf ExLVwU.View <> ViewConstants.vTiles Then
          ExLVwU.GetIconSpacing ViewConstants.vIcons, spacingLarge.cx, spacingLarge.cy
          ExLVwU.GetIconSpacing ViewConstants.vSmallIcons, spacingSmall.cx, spacingSmall.cy
        End If
        ExLVwU.GetOrigin origin.x, origin.y
        x = x + origin.x
        y = y + origin.y
      End If
      For i = LBound(files) To UBound(files)
        str = files(i)
        p = InStrRev(str, "\")
        str = Mid$(str, p + 1)
        Set itm = ExLVwU.ListItems.Add(str, , 1)
        If hasPositions Then
          If ExLVwU.View = ViewConstants.vTiles Then
            ' TODO: Find out how Windows is positioning the items
            itm.SetPosition 2.4 * positions(i).x + x, positions(i).y + y
          Else
            itm.SetPosition ((positions(i).x * spacingSmall.cx) / spacingLarge.cx) + x, ((positions(i).y * spacingSmall.cy) / spacingLarge.cy) + y
          End If
        End If
      Next i
NoFiles:
    End If
  End If

  lastDropTarget = -1
  If bComctl32Version600OrNewer And Not freeFloat Then
    ExLVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  End If

  ' Be careful with odeMove!! Some sources delete the data themselves if the Move effect was
  ' chosen. So you may lose data!
  'If shift And vbShiftMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeLink

'  If Data.GetFormat(vbCFText) Then MsgBox "Dropped Text: " & Data.GetData(vbCFText)
'  If Data.GetFormat(CF_OEMTEXT) Then MsgBox "Dropped OEM-Text: " & Data.GetData(7)
'  If Data.GetFormat(CF_UNICODETEXT) Then MsgBox "Dropped Unicode-Text: " & Data.GetData(13)
'  If Data.GetFormat(vbCFRTF) Then MsgBox "Dropped RTF-Text: " & Data.GetData(vbCFRTF)
'  If Data.GetFormat(CFSTR_HTML) Then MsgBox "Dropped HTML: " & Data.GetData(CFSTR_HTML)
'  If Data.GetFormat(CFSTR_HTML2) Then MsgBox "Dropped HTML2: " & Data.GetData(CFSTR_HTML2)
'  If Data.GetFormat(CFSTR_MOZILLAHTMLCONTEXT) Then MsgBox "Dropped text/_moz_htmlcontext: " & Data.GetData(CFSTR_MOZILLAHTMLCONTEXT)
'  If Data.GetFormat(CFSTR_TEXTHTML) Then MsgBox "Dropped text/html: " & Data.GetData(CFSTR_TEXTHTML)
'  If data.GetFormat(vbCFFiles) Then
'    str = ""
'    For Each file In data.Files
'      str = str & file & vbNewLine
'    Next file
'    MsgBox "Dropped files:" & vbNewLine & str
'  End If
'
'  ReDim bytArray(1 To LenB(CLSID_RecycleBin)) As Byte
'  CopyMemory VarPtr(bytArray(LBound(bytArray))), VarPtr(CLSID_RecycleBin), LenB(CLSID_RecycleBin)
'  Data.SetData CFSTR_TARGETCLSID, bytArray
'
'  ' another way to inform the source of the performed drop effect
'  Data.SetData CFSTR_PERFORMEDDROPEFFECT, Effect
End Sub

Private Sub ExLVwU_OLEDragEnter(ByVal Data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim dropActionDescription As String
  Dim dropDescriptionIcon As DropDescriptionIconConstants
  Dim dropTargetDescription As String
  Dim insertMarkPosition As InsertMarkPositionConstants
  Dim itm As ListViewItem
  Dim newDropTarget As Long

  dropDescriptionIcon = DropDescriptionIconConstants.ddiUseDefault
  If Data.GetFormat(vbCFDIB) Then
    effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
    dropActionDescription = "Use as background image"
    dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
  Else
    If dropTarget Is Nothing Then
      newDropTarget = -1
    Else
      newDropTarget = dropTarget.Index
    End If
    lastDropTarget = newDropTarget

    If bComctl32Version600OrNewer And Not freeFloat Then
      ExLVwU.GetClosestInsertMarkPosition Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), insertMarkPosition, itm
      ExLVwU.SetInsertMarkPosition insertMarkPosition, itm
      If Not (itm Is Nothing) Then
        dropTargetDescription = itm.Text
      End If
    End If

    effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
    If shift And vbShiftMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeMove
    If shift And vbCtrlMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
    If shift And vbAltMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeLink

    If itm Is Nothing Then
      If freeFloat Then
        Select Case effect
          Case ExLVwLibUCtl.OLEDropEffectConstants.odeMove
            dropActionDescription = "Move here"
            dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
          Case ExLVwLibUCtl.OLEDropEffectConstants.odeLink
            dropActionDescription = "Create link here"
            dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
          Case Else
            dropActionDescription = "Copy here"
            dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
        End Select
      Else
        dropDescriptionIcon = DropDescriptionIconConstants.ddiNoDrop
        dropActionDescription = "Cannot place here"
      End If
    Else
      Select Case effect
        Case ExLVwLibUCtl.OLEDropEffectConstants.odeMove
          Select Case insertMarkPosition
            Case InsertMarkPositionConstants.impAfter
              dropActionDescription = "Move after %1"
            Case InsertMarkPositionConstants.impBefore
              dropActionDescription = "Move before %1"
            Case Else
              dropActionDescription = "Move to %1"
          End Select
          dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
        Case ExLVwLibUCtl.OLEDropEffectConstants.odeLink
          Select Case insertMarkPosition
            Case InsertMarkPositionConstants.impAfter
              dropActionDescription = "Create link after %1"
            Case InsertMarkPositionConstants.impBefore
              dropActionDescription = "Create link before %1"
            Case Else
              dropActionDescription = "Create link in %1"
          End Select
          dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
        Case Else
          Select Case insertMarkPosition
            Case InsertMarkPositionConstants.impAfter
              dropActionDescription = "Insert copy after %1"
            Case InsertMarkPositionConstants.impBefore
              dropActionDescription = "Insert copy before %1"
            Case Else
              dropActionDescription = "Copy to %1"
          End Select
          dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
      End Select
    End If
  End If
  If dropDescriptionIcon = DropDescriptionIconConstants.ddiUseDefault Then
    Data.SetDropDescription , , dropDescriptionIcon
  Else
    Data.SetDropDescription dropTargetDescription, dropActionDescription, dropDescriptionIcon
  End If
End Sub

Private Sub ExLVwU_OLEDragLeave(ByVal Data As ExLVwLibUCtl.IOLEDataObject, ByVal dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
  lastDropTarget = -1
  If bComctl32Version600OrNewer And Not freeFloat Then
    ExLVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  End If
  Data.SetDropDescription , , DropDescriptionIconConstants.ddiUseDefault
End Sub

Private Sub ExLVwU_OLEDragMouseMove(ByVal Data As ExLVwLibUCtl.IOLEDataObject, effect As ExLVwLibUCtl.OLEDropEffectConstants, dropTarget As ExLVwLibUCtl.IListViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim dropActionDescription As String
  Dim dropDescriptionIcon As DropDescriptionIconConstants
  Dim dropTargetDescription As String
  Dim insertMarkPosition As InsertMarkPositionConstants
  Dim itm As ListViewItem
  Dim newDropTarget As Long

  dropDescriptionIcon = DropDescriptionIconConstants.ddiUseDefault
  If Data.GetFormat(vbCFDIB) Then
    effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
    dropActionDescription = "Use as background image"
    dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
  Else
    If dropTarget Is Nothing Then
      newDropTarget = -1
    Else
      newDropTarget = dropTarget.Index
    End If
    lastDropTarget = newDropTarget

    If bComctl32Version600OrNewer And Not freeFloat Then
      ExLVwU.GetClosestInsertMarkPosition Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), insertMarkPosition, itm
      ExLVwU.SetInsertMarkPosition insertMarkPosition, itm
      If Not (itm Is Nothing) Then
        dropTargetDescription = itm.Text
      End If
    End If

    effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
    If shift And vbShiftMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeMove
    If shift And vbCtrlMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeCopy
    If shift And vbAltMask Then effect = ExLVwLibUCtl.OLEDropEffectConstants.odeLink

    If itm Is Nothing Then
      If freeFloat Then
        Select Case effect
          Case ExLVwLibUCtl.OLEDropEffectConstants.odeMove
            dropActionDescription = "Move here"
            dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
          Case ExLVwLibUCtl.OLEDropEffectConstants.odeLink
            dropActionDescription = "Create link here"
            dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
          Case Else
            dropActionDescription = "Copy here"
            dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
        End Select
      Else
        dropDescriptionIcon = DropDescriptionIconConstants.ddiNoDrop
        dropActionDescription = "Cannot place here"
      End If
    Else
      Select Case effect
        Case ExLVwLibUCtl.OLEDropEffectConstants.odeMove
          Select Case insertMarkPosition
            Case InsertMarkPositionConstants.impAfter
              dropActionDescription = "Move after %1"
            Case InsertMarkPositionConstants.impBefore
              dropActionDescription = "Move before %1"
            Case Else
              dropActionDescription = "Move to %1"
          End Select
          dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
        Case ExLVwLibUCtl.OLEDropEffectConstants.odeLink
          Select Case insertMarkPosition
            Case InsertMarkPositionConstants.impAfter
              dropActionDescription = "Create link after %1"
            Case InsertMarkPositionConstants.impBefore
              dropActionDescription = "Create link before %1"
            Case Else
              dropActionDescription = "Create link in %1"
          End Select
          dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
        Case Else
          Select Case insertMarkPosition
            Case InsertMarkPositionConstants.impAfter
              dropActionDescription = "Insert copy after %1"
            Case InsertMarkPositionConstants.impBefore
              dropActionDescription = "Insert copy before %1"
            Case Else
              dropActionDescription = "Copy to %1"
          End Select
          dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
      End Select
    End If
  End If
  If dropDescriptionIcon = DropDescriptionIconConstants.ddiUseDefault Then
    Data.SetDropDescription , , dropDescriptionIcon
  Else
    Data.SetDropDescription dropTargetDescription, dropActionDescription, dropDescriptionIcon
  End If
End Sub

Private Sub ExLVwU_OLESetData(ByVal Data As ExLVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim itm As ListViewItem
  Dim str As String

  Select Case formatID
    Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
      For Each itm In ExLVwU.DraggedItems
        str = str & itm.Text & vbNewLine
      Next itm
      Data.SetData formatID, str

    Case vbCFFiles
      ' WARNING! Dragging files to objects like the recycle bin may result in deleted files and
      ' lost data!
'      ReDim files(1 To 2) As String
'      files(1) = "C:\Test1.txt"
'      files(2) = "C:\Test2.txt"
'      Data.SetData formatID, files

    Case vbCFBitmap, vbCFDIB, CF_DIBV5
      Data.SetData formatID, draggedPicture

    Case vbCFPalette
      If Not (draggedPicture Is Nothing) Then
        Data.SetData formatID, draggedPicture.hPal
      End If
  End Select
End Sub

Private Sub ExLVwU_OLEStartDrag(ByVal Data As ExLVwLibUCtl.IOLEDataObject)
  Data.SetData vbCFText
  Data.SetData CF_OEMTEXT
  Data.SetData CF_UNICODETEXT
  Data.SetData vbCFFiles
  If Not (draggedPicture Is Nothing) Then
    Data.SetData vbCFPalette
    Data.SetData vbCFBitmap
    Data.SetData vbCFDIB
    Data.SetData CF_DIBV5
  End If
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

  CFSTR_HTML = RegisterClipboardFormat(StrPtr("HTML Format"))
  CFSTR_HTML2 = RegisterClipboardFormat(StrPtr("HTML (HyperText Markup Language)"))
  CFSTR_LOGICALPERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Logical Performed DropEffect"))
  CFSTR_MOZILLAHTMLCONTEXT = RegisterClipboardFormat(StrPtr("text/_moz_htmlcontext"))
  CFSTR_PERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Performed DropEffect"))
  CFSTR_SHELLIDLISTOFFSET = RegisterClipboardFormat(StrPtr("Shell Object Offsets"))
  CFSTR_TARGETCLSID = RegisterClipboardFormat(StrPtr("TargetCLSID"))
  CFSTR_TEXTHTML = RegisterClipboardFormat(StrPtr("text/html"))

  CLSIDFromString StrPtr(strCLSID_RecycleBin), CLSID_RecycleBin

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
  Set ToolTip = New clsToolTip
  With ToolTip
    .Create Me.hWnd
    .Activate
    .AddTool chkOLEDragDrop.hWnd, "Check to start OLE Drag'n'Drop operations instead of normal ones.", , , False
  End With

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
  Set ToolTip = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub mnuAeroDragImages_Click()
  ExLVwU.OLEDragImageStyle = IIf(Not mnuAeroDragImages.Checked, OLEDragImageStyleConstants.odistAeroIfAvailable, OLEDragImageStyleConstants.odistClassic)
  mnuAeroDragImages.Checked = (ExLVwU.OLEDragImageStyle = OLEDragImageStyleConstants.odistAeroIfAvailable)
End Sub

Private Sub mnuAutoArrange_Click()
  ExLVwU.AutoArrangeItems = IIf(Not mnuAutoArrange.Checked, AutoArrangeItemsConstants.aaiNormal, AutoArrangeItemsConstants.aaiNone)
  mnuAutoArrange.Checked = (ExLVwU.AutoArrangeItems = AutoArrangeItemsConstants.aaiNormal)
  freeFloat = Not ((ExLVwU.AutoArrangeItems = AutoArrangeItemsConstants.aaiNormal) And (ExLVwU.View <> ViewConstants.vDetails) And (ExLVwU.View <> ViewConstants.vList))
End Sub

Private Sub mnuCancel_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCancel
End Sub

Private Sub mnuCopy_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCopy
End Sub

Private Sub mnuItemAlignmentLeft_Click()
  ExLVwU.ItemAlignment = iaLeft
  UpdateMenu
End Sub

Private Sub mnuItemAlignmentTop_Click()
  ExLVwU.ItemAlignment = iaTop
  UpdateMenu
End Sub

Private Sub mnuMove_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciMove
End Sub

Private Sub mnuSnapToGrid_Click()
  ExLVwU.SnapToGrid = Not mnuSnapToGrid.Checked
  mnuSnapToGrid.Checked = ExLVwU.SnapToGrid
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
    .Add "Column 1", , 100, , , 1
    .Add "Column 2", , 160, , alCenter, 2
    .Add "Column 3", , 160, , alRight, 3
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
    Set itm = .Add("Item 1", , iIcon, , , , Array(si1a, si2a))
    itm.SubItems(1).Text = "Item 1, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 1, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 2", , iIcon, , , , Array(si2b, si1b))
    itm.SubItems(1).Text = "Item 2, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 2, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 3", , iIcon, , , , Array(si1a, si2a))
    itm.SubItems(1).Text = "Item 3, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 3, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 4", , iIcon, , , , Array(si2b, si1b))
    itm.SubItems(1).Text = "Item 4, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 4, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 5", , iIcon, , , , Array(si1a, si2a))
    itm.SubItems(1).Text = "Item 5, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 5, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 6", , iIcon, , , , Array(si2b, si1b))
    itm.SubItems(1).Text = "Item 6, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 6, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 7", , iIcon, , , , Array(si1a, si2a))
    itm.SubItems(1).Text = "Item 7, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 7, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 8", , iIcon, , , , Array(si2b, si1b))
    itm.SubItems(1).Text = "Item 8, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 8, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 9", , iIcon, , , , Array(si1a, si2a))
    itm.SubItems(1).Text = "Item 9, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 9, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
    iIcon = (iIcon + 1) Mod cImages

    Set itm = .Add("Item 10", , iIcon, , , , Array(si2b, si1b))
    itm.SubItems(1).Text = "Item 10, SubItem 1"
    itm.SubItems(1).IconIndex = (iIcon + 1) Mod cImages
    itm.SubItems(2).Text = "Item 10, SubItem 2"
    itm.SubItems(2).IconIndex = (iIcon + 2) Mod cImages
  End With
End Sub

Private Function IsEqualGUID(IID1 As GUID, IID2 As GUID) As Boolean
  Dim Tmp1 As Currency
  Dim Tmp2 As Currency

  If IID1.Data1 = IID2.Data1 Then
    If IID1.Data2 = IID2.Data2 Then
      If IID1.Data3 = IID2.Data3 Then
        CopyMemory VarPtr(Tmp1), VarPtr(IID1.Data4(0)), 8
        CopyMemory VarPtr(Tmp2), VarPtr(IID2.Data4(0)), 8
        IsEqualGUID = (Tmp1 = Tmp2)
      End If
    End If
  End If
End Function

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
  mnuAutoArrange.Checked = (ExLVwU.AutoArrangeItems = AutoArrangeItemsConstants.aaiNormal)
  mnuSnapToGrid.Checked = ExLVwU.SnapToGrid
  Select Case ExLVwU.ItemAlignment
    Case iaTop
      mnuItemAlignmentTop.Checked = True
      mnuItemAlignmentLeft.Checked = False
    Case iaLeft
      mnuItemAlignmentTop.Checked = False
      mnuItemAlignmentLeft.Checked = True
  End Select
  freeFloat = Not ((ExLVwU.AutoArrangeItems = AutoArrangeItemsConstants.aaiNormal) And (ExLVwU.View <> ViewConstants.vDetails) And (ExLVwU.View <> ViewConstants.vList))
End Sub
