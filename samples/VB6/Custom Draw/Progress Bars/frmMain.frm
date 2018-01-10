VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.7#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerListView 1.7 - CustomDraw sample"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
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
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   240
      Top             =   3720
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLvw 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      _cx             =   10610
      _cy             =   6165
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
      ClickableColumnHeaders=   0   'False
      ColumnHeaderVisibility=   1
      DisabledEvents  =   1048319
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
      Left            =   2460
      TabIndex        =   0
      Top             =   3720
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


  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private hBrush As Long
  Private percentage(0 To 4) As Double
  Private themeableOS As Boolean


  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
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
  ExLvw.About
End Sub

Private Sub ExLvw_CreatedHeaderControlWindow(ByVal hWndHeader As Long)
  InsertColumns
End Sub

Private Sub ExLvw_CustomDraw(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal drawAllItems As Boolean, textColor As stdole.OLE_COLOR, TextBackColor As stdole.OLE_COLOR, ByVal drawStage As ExLVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExLVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExLVwLibUCtl.RECTANGLE, furtherProcessing As ExLVwLibUCtl.CustomDrawReturnValuesConstants)
  Dim rc As RECT

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case CustomDrawStageConstants.cdsItemPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault Or CustomDrawReturnValuesConstants.cdrvNotifySubItemDraw

    Case CustomDrawStageConstants.cdsSubItemPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault Or CustomDrawReturnValuesConstants.cdrvNotifyPostPaint

    Case CustomDrawStageConstants.cdsSubItemPostPaint
      If listSubItem.Index = 1 Then
        listSubItem.GetRectangle ItemRectangleTypeConstants.irtEntireItem, , rc.Top, , rc.Bottom
        rc.Left = drawingRectangle.Left
        rc.Top = rc.Top + 1
        rc.Bottom = rc.Bottom - 2
        rc.Right = drawingRectangle.Right - 55
        If rc.Right > rc.Left Then
          rc.Right = rc.Left + percentage(listItem.Index) * (rc.Right - rc.Left)

          FillRect hDC, rc, hBrush
        End If
      Else
        furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault
      End If
  End Select
End Sub

Private Sub ExLvw_ItemGetDisplayInfo(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal requestedInfo As ExLVwLibUCtl.RequestedInfoConstants, IconIndex As Long, Indent As Long, groupID As Long, TileViewColumns() As ExLVwLibUCtl.TILEVIEWSUBITEM, ByVal maxItemTextLength As Long, itemText As String, OverlayIndex As Long, StateImageIndex As Long, itemStates As ExLVwLibUCtl.ItemStateConstants, dontAskAgain As Boolean)
  If requestedInfo And RequestedInfoConstants.riItemText Then
    If Not (listSubItem Is Nothing) Then
      itemText = Format(percentage(listItem.Index), "0.0 %")
    End If
  End If
End Sub

Private Sub ExLvw_RecreatedControlWindow(ByVal hWnd As Long)
  InsertItems
End Sub

Private Sub Form_Initialize()
  Const COLOR_HIGHLIGHT = 13
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  percentage(0) = 0
  percentage(1) = 0.1
  percentage(2) = 0.2
  percentage(3) = 0.3
  percentage(4) = 0.4

  hBrush = GetSysColorBrush(COLOR_HIGHLIGHT)
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  InsertColumns
  InsertItems
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub Timer1_Timer()
  Dim i As Long

  For i = 0 To 4
    percentage(i) = percentage(i) + 0.005
    If percentage(i) > 1 Then percentage(i) = 0
  Next i
  ExLvw.RedrawItems 0, 4
End Sub


Private Sub InsertColumns()
  With ExLvw.Columns
    .Add "Items", , 100
    .Add "Progress", , 250, , ExLVwLibUCtl.AlignmentConstants.alRight
  End With
End Sub

Private Sub InsertItems()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLvw.hWnd, StrPtr("explorer"), 0
  End If

  With ExLvw.ListItems
    .Add "Item 1"
    .Add "Item 2"
    .Add "Item 3"
    .Add "Item 4"
    .Add "Item 5"
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
