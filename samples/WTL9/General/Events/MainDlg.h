// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////
#pragma once
#include <initguid.h>

#import <libid:9FC6639B-4237-4fb5-93B8-24049D39DF74> version("1.7") raw_dispinterfaces
#import <libid:2B7096FB-1E68-4e54-A2D9-4AA933EAC680> version("1.7") raw_dispinterfaces

DEFINE_GUID(LIBID_ExLVwLibU, 0x9fc6639b, 0x4237, 0x4fb5, 0x93, 0xb8, 0x24, 0x04, 0x9d, 0x39, 0xdf, 0x74);
DEFINE_GUID(LIBID_ExLVwLibA, 0x2b7096fb, 0x1e68, 0x4e54, 0xa2, 0xd9, 0x4a, 0xa9, 0x33, 0xea, 0xc6, 0x80);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_EXLVWU, CMainDlg, &__uuidof(ExLVwLibU::_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>,
    public IDispEventImpl<IDC_EXLVWA, CMainDlg, &__uuidof(ExLVwLibA::_IExplorerListViewEvents), &LIBID_ExLVwLibA, 1, 7>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> exlvwUContainerWnd;
	CContainedWindowT<CWindow> exlvwUWnd;
	CContainedWindowT<CAxWindow> exlvwAContainerWnd;
	CContainedWindowT<CWindow> exlvwAWnd;

	CMainDlg() :
	    exlvwUContainerWnd(this, 1),
	    exlvwAContainerWnd(this, 2),
	    exlvwUWnd(this, 3),
	    exlvwAWnd(this, 4)
	{
	}

	BOOL exlvwaIsFocused;
	OLE_HANDLE hBMPDownArrow;
	OLE_HANDLE hBMPUpArrow;
	CComQIPtr<IPicture> pDraggedPicture;
	POINT ptStartDrag;

	struct Controls
	{
		CEdit logEdit;
		CButton aboutButton;
		CImageList smallImageList;
		CImageList largeImageList;
		CImageList extralargeImageList;
		CComPtr<ExLVwLibU::IExplorerListView> exlvwU;
		CComPtr<ExLVwLibA::IExplorerListView> exlvwA;

		~Controls()
		{
			if(!smallImageList.IsNull()) {
				smallImageList.Destroy();
			}
			if(!largeImageList.IsNull()) {
				largeImageList.Destroy();
			}
			if(!extralargeImageList.IsNull()) {
				extralargeImageList.Destroy();
			}
		}
	} controls;

	BOOL IsComctl32Version600OrNewer(void);
	BOOL IsComctl32Version610OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(ID_VIEW_ICONS, OnViewIcons)
		COMMAND_ID_HANDLER(ID_VIEW_SMALLICONS, OnViewSmallIcons)
		COMMAND_ID_HANDLER(ID_VIEW_LIST, OnViewList)
		COMMAND_ID_HANDLER(ID_VIEW_DETAILS, OnViewDetails)
		COMMAND_ID_HANDLER(ID_VIEW_TILES, OnViewTiles)
		COMMAND_ID_HANDLER(ID_VIEW_EXTENDEDTILES, OnViewExtendedTiles)
		COMMAND_ID_HANDLER(ID_VIEW_AUTOARRANGE, OnViewAutoArrange)
		COMMAND_ID_HANDLER(ID_VIEW_JUSTIFYICONCOLUMNS, OnViewJustifyIconColumns)
		COMMAND_ID_HANDLER(ID_VIEW_SNAPTOGRID, OnViewSnapToGrid)
		COMMAND_ID_HANDLER(ID_VIEW_SHOWINGROUPS, OnViewShowInGroups)
		COMMAND_ID_HANDLER(ID_VIEW_CHECKONSELECT, OnViewCheckOnSelect)
		COMMAND_ID_HANDLER(ID_VIEW_AUTOSIZECOLUMNS, OnViewAutoSizeColumns)
		COMMAND_ID_HANDLER(ID_VIEW_RESIZABLECOLUMNS, OnViewResizableColumns)
		COMMAND_ID_HANDLER(ID_ALIGN_TOP, OnAlignTop)
		COMMAND_ID_HANDLER(ID_ALIGN_LEFT, OnAlignLeft)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusU)

		ALT_MSG_MAP(4)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusA)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 1, AbortedDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 2, AfterScrollExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 3, BeforeScrollExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 4, BeginColumnResizingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 5, BeginMarqueeSelectionExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 6, CacheItemsHintExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 0, CaretChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 7, ChangedWorkAreasExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 8, ChangingWorkAreasExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 9, ClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 10, ColumnBeginDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 11, ColumnEndAutoDragDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 12, ColumnMouseEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 13, ColumnMouseLeaveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 14, CompareGroupsExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 15, CompareItemsExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 16, ContextMenuExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 17, CreatedEditControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 19, CustomDrawExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 20, DblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 21, DestroyedControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 22, DestroyedEditControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 23, DestroyedHeaderControlWindowExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 24, DragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 25, DropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 26, EditChangeExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 27, EditClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 28, EditContextMenuExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 29, EditDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 30, EditGotFocusExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 31, EditKeyDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 32, EditKeyPressExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 33, EditKeyUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 34, EditLostFocusExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 35, EditMClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 36, EditMDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 37, EditMouseDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 38, EditMouseEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 39, EditMouseHoverExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 40, EditMouseLeaveExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 41, EditMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 42, EditMouseUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 43, EditRClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 44, EditRDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 45, EndColumnResizingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 46, FilterButtonClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 47, FilterChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 48, FindVirtualItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 49, FreeColumnDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 50, FreeItemDataExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 51, GroupCustomDrawExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 52, HeaderAbortedDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 53, HeaderClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 54, HeaderContextMenuExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 55, HeaderCustomDrawExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 56, HeaderDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 57, HeaderDividerDblClickExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 58, HeaderDragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 59, HeaderDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 60, HeaderItemGetDisplayInfoExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 61, HeaderMClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 62, HeaderMDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 63, HeaderMouseDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 64, HeaderMouseEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 65, HeaderMouseHoverExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 66, HeaderMouseLeaveExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 67, HeaderMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 68, HeaderMouseUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 69, HeaderOLECompleteDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 70, HeaderOLEDragDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 71, HeaderOLEDragEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 72, HeaderOLEDragLeaveExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 73, HeaderOLEDragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 74, HeaderOLEGiveFeedbackExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 75, HeaderOLEQueryContinueDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 76, HeaderOLESetDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 77, HeaderOLEStartDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 78, HeaderOwnerDrawItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 79, HeaderRClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 80, HeaderRDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 81, HotItemChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 82, HotItemChangingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 83, IncrementalSearchStringChangingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 84, InsertedColumnExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 85, InsertedGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 86, InsertedItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 87, InsertingColumnExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 88, InsertingGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 89, InsertingItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 90, ItemActivateExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 91, ItemBeginDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 92, ItemBeginRDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 93, ItemGetDisplayInfoExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 94, ItemGetInfoTipTextExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 95, ItemMouseEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 96, ItemMouseLeaveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 97, ItemSelectionChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 98, ItemSetTextExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 99, ItemStateImageChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 100, ItemStateImageChangingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 101, KeyDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 102, KeyPressExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 103, KeyUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 104, MClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 105, MDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 106, MouseDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 107, MouseEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 108, MouseHoverExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 109, MouseLeaveExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 110, MouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 111, MouseUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 112, OLECompleteDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 113, OLEDragDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 114, OLEDragEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 115, OLEDragLeaveExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 116, OLEDragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 117, OLEGiveFeedbackExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 118, OLEQueryContinueDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 119, OLESetDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 120, OLEStartDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 121, OwnerDrawItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 122, RClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 123, RDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 124, RecreatedControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 125, RemovedColumnExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 126, RemovedGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 127, RemovedItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 128, RemovingColumnExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 129, RemovingGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 130, RemovingItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 131, RenamedItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 132, RenamingItemExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 133, ResizedControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 134, ResizingColumnExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 135, SelectedItemRangeExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 136, StartingLabelEditingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 137, SubItemMouseEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 138, SubItemMouseLeaveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 139, IncrementalSearchingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 140, ChangedSortOrderExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 141, ChangingSortOrderExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 142, ColumnClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 143, ColumnDropDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 144, ColumnStateImageChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 145, ColumnStateImageChangingExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 146, EmptyMarkupTextLinkClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 147, GroupAsynchronousDrawFailedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 148, GroupTaskLinkClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 149, HeaderChevronClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 150, HeaderGotFocusExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 151, HeaderKeyDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 152, HeaderKeyPressExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 153, HeaderKeyUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 154, HeaderLostFocusExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 155, HeaderOLEDragEnterPotentialTargetExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 156, HeaderOLEDragLeavePotentialTargetExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 157, HeaderOLEReceivedNewDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 158, ItemAsynchronousDrawFailedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 159, OLEDragEnterPotentialTargetExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 160, OLEDragLeavePotentialTargetExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 161, OLEReceivedNewDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 162, FooterItemClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 163, FreeFooterItemDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 164, ItemGetGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 165, ItemGetOccurrencesCountExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 166, MapGroupWideToTotalItemIndexExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 167, SettingItemInfoTipTextExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 168, CollapsedGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 169, ExpandedGroupExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 170, GroupGotFocusExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 171, GroupLostFocusExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 172, GroupSelectionChangedExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 173, CancelSubItemEditExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 174, ConfigureSubItemControlExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 175, EndSubItemEditExlvwu)
		//SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 176, GetSubItemControlExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 177, InvokeVerbFromSubItemControlExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 178, EditMouseWheelExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 179, EditXClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 180, EditXDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 181, HeaderMouseWheelExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 182, HeaderXClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 183, HeaderXDblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 184, MouseWheelExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 185, XClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(ExLVwLibU::_IExplorerListViewEvents), 186, XDblClickExlvwu)

		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 1, AbortedDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 2, AfterScrollExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 3, BeforeScrollExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 4, BeginColumnResizingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 5, BeginMarqueeSelectionExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 6, CacheItemsHintExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 0, CaretChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 7, ChangedWorkAreasExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 8, ChangingWorkAreasExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 9, ClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 10, ColumnBeginDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 11, ColumnEndAutoDragDropExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 12, ColumnMouseEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 13, ColumnMouseLeaveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 14, CompareGroupsExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 15, CompareItemsExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 16, ContextMenuExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 17, CreatedEditControlWindowExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 19, CustomDrawExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 20, DblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 21, DestroyedControlWindowExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 22, DestroyedEditControlWindowExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 23, DestroyedHeaderControlWindowExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 24, DragMouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 25, DropExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 26, EditChangeExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 27, EditClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 28, EditContextMenuExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 29, EditDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 30, EditGotFocusExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 31, EditKeyDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 32, EditKeyPressExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 33, EditKeyUpExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 34, EditLostFocusExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 35, EditMClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 36, EditMDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 37, EditMouseDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 38, EditMouseEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 39, EditMouseHoverExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 40, EditMouseLeaveExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 41, EditMouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 42, EditMouseUpExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 43, EditRClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 44, EditRDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 45, EndColumnResizingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 46, FilterButtonClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 47, FilterChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 48, FindVirtualItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 49, FreeColumnDataExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 50, FreeItemDataExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 51, GroupCustomDrawExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 52, HeaderAbortedDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 53, HeaderClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 54, HeaderContextMenuExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 55, HeaderCustomDrawExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 56, HeaderDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 57, HeaderDividerDblClickExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 58, HeaderDragMouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 59, HeaderDropExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 60, HeaderItemGetDisplayInfoExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 61, HeaderMClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 62, HeaderMDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 63, HeaderMouseDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 64, HeaderMouseEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 65, HeaderMouseHoverExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 66, HeaderMouseLeaveExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 67, HeaderMouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 68, HeaderMouseUpExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 69, HeaderOLECompleteDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 70, HeaderOLEDragDropExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 71, HeaderOLEDragEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 72, HeaderOLEDragLeaveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 73, HeaderOLEDragMouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 74, HeaderOLEGiveFeedbackExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 75, HeaderOLEQueryContinueDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 76, HeaderOLESetDataExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 77, HeaderOLEStartDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 78, HeaderOwnerDrawItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 79, HeaderRClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 80, HeaderRDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 81, HotItemChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 82, HotItemChangingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 83, IncrementalSearchStringChangingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 84, InsertedColumnExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 85, InsertedGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 86, InsertedItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 87, InsertingColumnExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 88, InsertingGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 89, InsertingItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 90, ItemActivateExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 91, ItemBeginDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 92, ItemBeginRDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 93, ItemGetDisplayInfoExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 94, ItemGetInfoTipTextExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 95, ItemMouseEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 96, ItemMouseLeaveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 97, ItemSelectionChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 98, ItemSetTextExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 99, ItemStateImageChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 100, ItemStateImageChangingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 101, KeyDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 102, KeyPressExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 103, KeyUpExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 104, MClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 105, MDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 106, MouseDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 107, MouseEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 108, MouseHoverExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 109, MouseLeaveExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 110, MouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 111, MouseUpExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 112, OLECompleteDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 113, OLEDragDropExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 114, OLEDragEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 115, OLEDragLeaveExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 116, OLEDragMouseMoveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 117, OLEGiveFeedbackExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 118, OLEQueryContinueDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 119, OLESetDataExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 120, OLEStartDragExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 121, OwnerDrawItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 122, RClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 123, RDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 124, RecreatedControlWindowExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 125, RemovedColumnExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 126, RemovedGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 127, RemovedItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 128, RemovingColumnExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 129, RemovingGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 130, RemovingItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 131, RenamedItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 132, RenamingItemExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 133, ResizedControlWindowExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 134, ResizingColumnExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 135, SelectedItemRangeExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 136, StartingLabelEditingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 137, SubItemMouseEnterExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 138, SubItemMouseLeaveExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 139, IncrementalSearchingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 140, ChangedSortOrderExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 141, ChangingSortOrderExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 142, ColumnClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 143, ColumnDropDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 144, ColumnStateImageChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 145, ColumnStateImageChangingExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 146, EmptyMarkupTextLinkClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 147, GroupAsynchronousDrawFailedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 148, GroupTaskLinkClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 149, HeaderChevronClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 150, HeaderGotFocusExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 151, HeaderKeyDownExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 152, HeaderKeyPressExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 153, HeaderKeyUpExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 154, HeaderLostFocusExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 155, HeaderOLEDragEnterPotentialTargetExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 156, HeaderOLEDragLeavePotentialTargetExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 157, HeaderOLEReceivedNewDataExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 158, ItemAsynchronousDrawFailedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 159, OLEDragEnterPotentialTargetExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 160, OLEDragLeavePotentialTargetExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 161, OLEReceivedNewDataExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 162, FooterItemClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 163, FreeFooterItemDataExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 164, ItemGetGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 165, ItemGetOccurrencesCountExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 166, MapGroupWideToTotalItemIndexExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 167, SettingItemInfoTipTextExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 168, CollapsedGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 169, ExpandedGroupExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 170, GroupGotFocusExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 171, GroupLostFocusExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 172, GroupSelectionChangedExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 173, CancelSubItemEditExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 174, ConfigureSubItemControlExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 175, EndSubItemEditExlvwa)
		//SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 176, GetSubItemControlExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 177, InvokeVerbFromSubItemControlExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 178, EditMouseWheelExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 179, EditXClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 180, EditXDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 181, HeaderMouseWheelExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 180, HeaderXClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 181, HeaderXDblClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 182, MouseWheelExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 183, XClickExlvwa)
		SINK_ENTRY_EX(IDC_EXLVWA, __uuidof(ExLVwLibA::_IExplorerListViewEvents), 184, XDblClickExlvwa)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewSmallIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewDetails(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewExtendedTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewAutoArrange(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewJustifyIconColumns(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewSnapToGrid(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewShowInGroups(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewCheckOnSelect(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewAutoSizeColumns(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewResizableColumns(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlignTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlignLeft(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);
	void InsertColumnsA(void);
	void InsertFooterItemsA(void);
	void InsertGroupsA(void);
	void InsertItemsA(void);
	void InsertColumnsU(void);
	void InsertFooterItemsU(void);
	void InsertGroupsU(void);
	void InsertItemsU(void);
	void UpdateMenu(void);

	void __stdcall AbortedDragExlvwu();
	void __stdcall AfterScrollExlvwu(long dx, long dy);
	void __stdcall BeforeScrollExlvwu(long dx, long dy);
	void __stdcall BeginColumnResizingExlvwu(LPDISPATCH column, VARIANT_BOOL* cancel);
	void __stdcall BeginMarqueeSelectionExlvwu(BOOL* cancel);
	void __stdcall CacheItemsHintExlvwu(LPDISPATCH firstItem, LPDISPATCH lastItem);
	void __stdcall CancelSubItemEditExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemEditModeConstants editMode);
	void __stdcall CaretChangedExlvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem);
	void __stdcall ChangedSortOrderExlvwu(long previousSortOrder, long newSortOrder);
	void __stdcall ChangedWorkAreasExlvwu(LPDISPATCH WorkAreas);
	void __stdcall ChangingSortOrderExlvwu(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange);
	void __stdcall ChangingWorkAreasExlvwu(LPDISPATCH WorkAreas, VARIANT_BOOL* cancelChanges);
	void __stdcall ClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall CollapsedGroupExlvwu(LPDISPATCH Group);
	void __stdcall ColumnBeginDragExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDragDrop);
	void __stdcall ColumnClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ColumnDropDownExlvwu(LPDISPATCH column, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall ColumnEndAutoDragDropExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDrop);
	void __stdcall ColumnMouseEnterExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ColumnMouseLeaveExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ColumnStateImageChangedExlvwu(LPDISPATCH column, long previousStateImageIndex, long newStateImageIndex, long causedBy);
	void __stdcall ColumnStateImageChangingExlvwu(LPDISPATCH column, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange);
	void __stdcall CompareGroupsExlvwu(LPDISPATCH firstGroup, LPDISPATCH secondGroup, long* result);
	void __stdcall CompareItemsExlvwu(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result);
	void __stdcall ConfigureSubItemControlExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemControlKindConstants controlKind, ExLVwLibU::SubItemEditModeConstants editMode, ExLVwLibU::SubItemControlConstants subItemControl, BSTR* themeAppName, BSTR* themeIDList, long* hFont, OLE_COLOR* textColor, long* pPropertyDescription, long pPropertyValue);
	void __stdcall ContextMenuExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedEditControlWindowExlvwu(long hWndEdit);
	void __stdcall CreatedHeaderControlWindowExlvwu(long hWndHeader);
	void __stdcall CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL drawAllItems, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExLVwLibU::CustomDrawStageConstants drawStage, ExLVwLibU::CustomDrawItemStateConstants itemState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle, ExLVwLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowExlvwu(long hWnd);
	void __stdcall DestroyedEditControlWindowExlvwu(long hWndEdit);
	void __stdcall DestroyedHeaderControlWindowExlvwu(long hWndHeader);
	void __stdcall DragMouseMoveExlvwu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropExlvwu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall EditChangeExlvwu();
	void __stdcall EditClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditContextMenuExlvwu(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall EditDblClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditGotFocusExlvwu();
	void __stdcall EditKeyDownExlvwu(short* keyCode, short shift);
	void __stdcall EditKeyPressExlvwu(short* keyAscii);
	void __stdcall EditKeyUpExlvwu(short* keyCode, short shift);
	void __stdcall EditLostFocusExlvwu();
	void __stdcall EditMClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMDblClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseDownExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseEnterExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseHoverExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseLeaveExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseMoveExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseUpExlvwu(short button, short shift, long x, long y);
	void __stdcall EditMouseWheelExlvwu(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall EditRClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditRDblClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditXClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EditXDblClickExlvwu(short button, short shift, long x, long y);
	void __stdcall EmptyMarkupTextLinkClickExlvwu(long linkIndex, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall EndColumnResizingExlvwu(LPDISPATCH column);
	void __stdcall EndSubItemEditExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemEditModeConstants editMode, long pPropertyKey, long pPropertyValue, VARIANT_BOOL* cancel);
	void __stdcall ExpandedGroupExlvwu(LPDISPATCH Group);
	void __stdcall FilterButtonClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, ExLVwLibU::RECTANGLE* filterButtonRectangle, VARIANT_BOOL* raiseFilterChanged);
	void __stdcall FilterChangedExlvwu(LPDISPATCH column);
	void __stdcall FindVirtualItemExlvwu(LPDISPATCH itemToStartWith, long searchMode, VARIANT* searchFor, long searchDirection, VARIANT_BOOL wrapAtLastItem, LPDISPATCH* foundItem);
	void __stdcall FooterItemClickExlvwu(LPDISPATCH footerItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* removeFooterArea);
	void __stdcall FreeColumnDataExlvwu(LPDISPATCH column);
	void __stdcall FreeFooterItemDataExlvwu(LPDISPATCH footerItem, LONG itemData);
	void __stdcall FreeItemDataExlvwu(LPDISPATCH listItem);
	void __stdcall GetSubItemControlExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemControlKindConstants controlKind, ExLVwLibU::SubItemEditModeConstants editMode, ExLVwLibU::SubItemControlConstants* subItemControl);
	void __stdcall GroupAsynchronousDrawFailedExlvwu(LPDISPATCH Group, ExLVwLibU::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibU::ImageDrawingFailureReasonConstants failureReason, ExLVwLibU::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw);
	void __stdcall GroupCustomDrawExlvwu(LPDISPATCH Group, OLE_COLOR* textColor, ExLVwLibU::AlignmentConstants* headerAlignment, ExLVwLibU::AlignmentConstants* footerAlignment, ExLVwLibU::CustomDrawStageConstants drawStage, ExLVwLibU::CustomDrawItemStateConstants groupState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle, ExLVwLibU::RECTANGLE* textRectangle, ExLVwLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall GroupGotFocusExlvwu(LPDISPATCH Group);
	void __stdcall GroupLostFocusExlvwu(LPDISPATCH Group);
	void __stdcall GroupSelectionChangedExlvwu(LPDISPATCH Group);
	void __stdcall GroupTaskLinkClickExlvwu(LPDISPATCH Group, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderAbortedDragExlvwu();
	void __stdcall HeaderChevronClickExlvwu(LPDISPATCH firstOverflownColumn, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall HeaderClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderContextMenuExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall HeaderCustomDrawExlvwu(LPDISPATCH column, ExLVwLibU::CustomDrawStageConstants drawStage, ExLVwLibU::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle, ExLVwLibU::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall HeaderDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderDividerDblClickExlvwu(LPDISPATCH column, VARIANT_BOOL* autoSizeColumn);
	void __stdcall HeaderDragMouseMoveExlvwu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall HeaderDropExlvwu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderGotFocusExlvwu();
	void __stdcall HeaderItemGetDisplayInfoExlvwu(LPDISPATCH column, long requestedInfo, long* IconIndex, long maxColumnCaptionLength, BSTR* columnCaption, VARIANT_BOOL* dontAskAgain);
	void __stdcall HeaderKeyDownExlvwu(short* keyCode, short shift);
	void __stdcall HeaderKeyPressExlvwu(short* keyAscii);
	void __stdcall HeaderKeyUpExlvwu(short* keyCode, short shift);
	void __stdcall HeaderLostFocusExlvwu();
	void __stdcall HeaderMClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseDownExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseEnterExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseHoverExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseLeaveExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseMoveExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseUpExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseWheelExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall HeaderOLECompleteDragExlvwu(LPDISPATCH data, long performedEffect);
	void __stdcall HeaderOLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails);
	void __stdcall HeaderOLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall HeaderOLEDragEnterPotentialTargetExlvwu(long hWndPotentialTarget);
	void __stdcall HeaderOLEDragLeaveExlvwu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails);
	void __stdcall HeaderOLEDragLeavePotentialTargetExlvwu(void);
	void __stdcall HeaderOLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall HeaderOLEGiveFeedbackExlvwu(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall HeaderOLEQueryContinueDragExlvwu(BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall HeaderOLEReceivedNewDataExlvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall HeaderOLESetDataExlvwu(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect);
	void __stdcall HeaderOLEStartDragExlvwu(LPDISPATCH data);
	void __stdcall HeaderOwnerDrawItemExlvwu(LPDISPATCH column, ExLVwLibU::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle);
	void __stdcall HeaderRClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderRDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderXClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderXDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HotItemChangedExlvwu(LPDISPATCH previousHotItem, LPDISPATCH newHotItem);
	void __stdcall HotItemChangingExlvwu(LPDISPATCH previousHotItem, LPDISPATCH newHotItem, VARIANT_BOOL* cancelChange);
	void __stdcall IncrementalSearchingExlvwu(BSTR currentSearchString, LONG* itemToSelect);
	void __stdcall IncrementalSearchStringChangingExlvwu(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange);
	void __stdcall InsertedColumnExlvwu(LPDISPATCH column);
	void __stdcall InsertedGroupExlvwu(LPDISPATCH Group);
	void __stdcall InsertedItemExlvwu(LPDISPATCH listItem);
	void __stdcall InsertingColumnExlvwu(LPDISPATCH column, VARIANT_BOOL* cancelInsertion);
	void __stdcall InsertingGroupExlvwu(LPDISPATCH Group, VARIANT_BOOL* cancelInsertion);
	void __stdcall InsertingItemExlvwu(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall InvokeVerbFromSubItemControlExlvwu(LPDISPATCH listItem, BSTR verb);
	void __stdcall ItemActivateExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short shift, long x, long y);
	void __stdcall ItemAsynchronousDrawFailedExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, ExLVwLibU::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibU::ImageDrawingFailureReasonConstants failureReason, ExLVwLibU::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw);
	void __stdcall ItemBeginDragExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemBeginRDragExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemGetDisplayInfoExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* Indent, long* groupID, LPSAFEARRAY* TileViewColumns, long maxItemTextLength, BSTR* itemText, long* OverlayIndex, long* StateImageIndex, long* itemStates, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemGetGroupExlvwu(LONG itemIndex, LONG occurrenceIndex, LONG* groupIndex);
	void __stdcall ItemGetInfoTipTextExlvwu(LPDISPATCH listItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall ItemGetOccurrencesCountExlvwu(LONG itemIndex, LONG* occurrencesCount);
	void __stdcall ItemMouseEnterExlvwu(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemMouseLeaveExlvwu(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemSelectionChangedExlvwu(LPDISPATCH listItem);
	void __stdcall ItemSetTextExlvwu(LPDISPATCH listItem, BSTR itemText);
	void __stdcall ItemStateImageChangedExlvwu(LPDISPATCH listItem, long previousStateImageIndex, long newStateImageIndex, long causedBy);
	void __stdcall ItemStateImageChangingExlvwu(LPDISPATCH listItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange);
	void __stdcall KeyDownExlvwu(short* keyCode, short shift);
	void __stdcall KeyPressExlvwu(short* keyAscii);
	void __stdcall KeyUpExlvwu(short* keyCode, short shift);
	void __stdcall MapGroupWideToTotalItemIndexExlvwu(LONG groupIndex, LONG groupWideItemIndex, LONG* totalItemIndex);
	void __stdcall MClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseWheelExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragExlvwu(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetExlvwu(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveExlvwu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetExlvwu(void);
	void __stdcall OLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackExlvwu(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragExlvwu(BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall OLEReceivedNewDataExlvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataExlvwu(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect);
	void __stdcall OLEStartDragExlvwu(LPDISPATCH data);
	void __stdcall OwnerDrawItemExlvwu(LPDISPATCH listItem, ExLVwLibU::OwnerDrawItemStateConstants itemState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle);
	void __stdcall RClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowExlvwu(long hWnd);
	void __stdcall RemovedColumnExlvwu(LPDISPATCH column);
	void __stdcall RemovedGroupExlvwu(LPDISPATCH Group);
	void __stdcall RemovedItemExlvwu(LPDISPATCH listItem);
	void __stdcall RemovingColumnExlvwu(LPDISPATCH column, VARIANT_BOOL* cancelDeletion);
	void __stdcall RemovingGroupExlvwu(LPDISPATCH Group, VARIANT_BOOL* cancelDeletion);
	void __stdcall RemovingItemExlvwu(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall RenamedItemExlvwu(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText);
	void __stdcall RenamingItemExlvwu(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming);
	void __stdcall ResizedControlWindowExlvwu();
	void __stdcall ResizingColumnExlvwu(LPDISPATCH column, long* newColumnWidth, VARIANT_BOOL* abortResizing);
	void __stdcall SelectedItemRangeExlvwu(LPDISPATCH firstItem, LPDISPATCH lastItem);
	void __stdcall SettingItemInfoTipTextExlvwu(LPDISPATCH listItem, BSTR* infoTipText, VARIANT_BOOL* abortInfoTip);
	void __stdcall StartingLabelEditingExlvwu(LPDISPATCH listItem, VARIANT_BOOL* cancelEditing);
	void __stdcall SubItemMouseEnterExlvwu(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall SubItemMouseLeaveExlvwu(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);

	void __stdcall AbortedDragExlvwa();
	void __stdcall AfterScrollExlvwa(long dx, long dy);
	void __stdcall BeforeScrollExlvwa(long dx, long dy);
	void __stdcall BeginColumnResizingExlvwa(LPDISPATCH column, VARIANT_BOOL* cancel);
	void __stdcall BeginMarqueeSelectionExlvwa(BOOL* cancel);
	void __stdcall CacheItemsHintExlvwa(LPDISPATCH firstItem, LPDISPATCH lastItem);
	void __stdcall CancelSubItemEditExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemEditModeConstants editMode);
	void __stdcall CaretChangedExlvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem);
	void __stdcall ChangedSortOrderExlvwa(long previousSortOrder, long newSortOrder);
	void __stdcall ChangedWorkAreasExlvwa(LPDISPATCH WorkAreas);
	void __stdcall ChangingSortOrderExlvwa(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange);
	void __stdcall ChangingWorkAreasExlvwa(LPDISPATCH WorkAreas, VARIANT_BOOL* cancelChanges);
	void __stdcall ClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall CollapsedGroupExlvwa(LPDISPATCH Group);
	void __stdcall ColumnBeginDragExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDragDrop);
	void __stdcall ColumnClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ColumnDropDownExlvwa(LPDISPATCH column, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall ColumnEndAutoDragDropExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDrop);
	void __stdcall ColumnMouseEnterExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ColumnMouseLeaveExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ColumnStateImageChangedExlvwa(LPDISPATCH column, long previousStateImageIndex, long newStateImageIndex, long causedBy);
	void __stdcall ColumnStateImageChangingExlvwa(LPDISPATCH column, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange);
	void __stdcall CompareGroupsExlvwa(LPDISPATCH firstGroup, LPDISPATCH secondGroup, long* result);
	void __stdcall CompareItemsExlvwa(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result);
	void __stdcall ConfigureSubItemControlExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemControlKindConstants controlKind, ExLVwLibA::SubItemEditModeConstants editMode, ExLVwLibA::SubItemControlConstants subItemControl, BSTR* themeAppName, BSTR* themeIDList, long* hFont, OLE_COLOR* textColor, long* pPropertyDescription, long pPropertyValue);
	void __stdcall ContextMenuExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall CreatedEditControlWindowExlvwa(long hWndEdit);
	void __stdcall CreatedHeaderControlWindowExlvwa(long hWndHeader);
	void __stdcall CustomDrawExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL drawAllItems, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExLVwLibA::CustomDrawStageConstants drawStage, ExLVwLibA::CustomDrawItemStateConstants itemState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle, ExLVwLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowExlvwa(long hWnd);
	void __stdcall DestroyedEditControlWindowExlvwa(long hWndEdit);
	void __stdcall DestroyedHeaderControlWindowExlvwa(long hWndHeader);
	void __stdcall DragMouseMoveExlvwa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropExlvwa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall EditChangeExlvwa();
	void __stdcall EditClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditContextMenuExlvwa(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall EditDblClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditGotFocusExlvwa();
	void __stdcall EditKeyDownExlvwa(short* keyCode, short shift);
	void __stdcall EditKeyPressExlvwa(short* keyAscii);
	void __stdcall EditKeyUpExlvwa(short* keyCode, short shift);
	void __stdcall EditLostFocusExlvwa();
	void __stdcall EditMClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMDblClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseDownExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseEnterExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseHoverExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseLeaveExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseMoveExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseUpExlvwa(short button, short shift, long x, long y);
	void __stdcall EditMouseWheelExlvwa(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall EditRClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditRDblClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditXClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EditXDblClickExlvwa(short button, short shift, long x, long y);
	void __stdcall EmptyMarkupTextLinkClickExlvwa(long linkIndex, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall EndColumnResizingExlvwa(LPDISPATCH column);
	void __stdcall EndSubItemEditExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemEditModeConstants editMode, long pPropertyKey, long pPropertyValue, VARIANT_BOOL* cancel);
	void __stdcall ExpandedGroupExlvwa(LPDISPATCH Group);
	void __stdcall FilterButtonClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, ExLVwLibA::RECTANGLE* filterButtonRectangle, VARIANT_BOOL* raiseFilterChanged);
	void __stdcall FilterChangedExlvwa(LPDISPATCH column);
	void __stdcall FindVirtualItemExlvwa(LPDISPATCH itemToStartWith, long searchMode, VARIANT* searchFor, long searchDirection, VARIANT_BOOL wrapAtLastItem, LPDISPATCH* foundItem);
	void __stdcall FooterItemClickExlvwa(LPDISPATCH footerItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* removeFooterArea);
	void __stdcall FreeColumnDataExlvwa(LPDISPATCH column);
	void __stdcall FreeFooterItemDataExlvwa(LPDISPATCH footerItem, LONG itemData);
	void __stdcall FreeItemDataExlvwa(LPDISPATCH listItem);
	void __stdcall GetSubItemControlExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemControlKindConstants controlKind, ExLVwLibA::SubItemEditModeConstants editMode, ExLVwLibA::SubItemControlConstants* subItemControl);
	void __stdcall GroupAsynchronousDrawFailedExlvwa(LPDISPATCH Group, ExLVwLibA::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibA::ImageDrawingFailureReasonConstants failureReason, ExLVwLibA::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw);
	void __stdcall GroupCustomDrawExlvwa(LPDISPATCH Group, OLE_COLOR* textColor, ExLVwLibA::AlignmentConstants* headerAlignment, ExLVwLibA::AlignmentConstants* footerAlignment, ExLVwLibA::CustomDrawStageConstants drawStage, ExLVwLibA::CustomDrawItemStateConstants groupState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle, ExLVwLibA::RECTANGLE* textRectangle, ExLVwLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall GroupGotFocusExlvwa(LPDISPATCH Group);
	void __stdcall GroupLostFocusExlvwa(LPDISPATCH Group);
	void __stdcall GroupSelectionChangedExlvwa(LPDISPATCH Group);
	void __stdcall GroupTaskLinkClickExlvwa(LPDISPATCH Group, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderAbortedDragExlvwa();
	void __stdcall HeaderChevronClickExlvwa(LPDISPATCH firstOverflownColumn, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall HeaderClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderContextMenuExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall HeaderCustomDrawExlvwa(LPDISPATCH column, ExLVwLibA::CustomDrawStageConstants drawStage, ExLVwLibA::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle, ExLVwLibA::CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall HeaderDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderDividerDblClickExlvwa(LPDISPATCH column, VARIANT_BOOL* autoSizeColumn);
	void __stdcall HeaderDragMouseMoveExlvwa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall HeaderDropExlvwa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderGotFocusExlvwa();
	void __stdcall HeaderItemGetDisplayInfoExlvwa(LPDISPATCH column, long requestedInfo, long* IconIndex, long maxColumnCaptionLength, BSTR* columnCaption, VARIANT_BOOL* dontAskAgain);
	void __stdcall HeaderKeyDownExlvwa(short* keyCode, short shift);
	void __stdcall HeaderKeyPressExlvwa(short* keyAscii);
	void __stdcall HeaderKeyUpExlvwa(short* keyCode, short shift);
	void __stdcall HeaderLostFocusExlvwa();
	void __stdcall HeaderMClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseDownExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseEnterExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseHoverExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseLeaveExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseMoveExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseUpExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderMouseWheelExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall HeaderOLECompleteDragExlvwa(LPDISPATCH data, long performedEffect);
	void __stdcall HeaderOLEDragDropExlvwa(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails);
	void __stdcall HeaderOLEDragEnterExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall HeaderOLEDragEnterPotentialTargetExlvwa(long hWndPotentialTarget);
	void __stdcall HeaderOLEDragLeaveExlvwa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails);
	void __stdcall HeaderOLEDragLeavePotentialTargetExlvwa(void);
	void __stdcall HeaderOLEDragMouseMoveExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall HeaderOLEGiveFeedbackExlvwa(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall HeaderOLEQueryContinueDragExlvwa(BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall HeaderOLEReceivedNewDataExlvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall HeaderOLESetDataExlvwa(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect);
	void __stdcall HeaderOLEStartDragExlvwa(LPDISPATCH data);
	void __stdcall HeaderOwnerDrawItemExlvwa(LPDISPATCH column, ExLVwLibA::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle);
	void __stdcall HeaderRClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderRDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderXClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HeaderXDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall HotItemChangedExlvwa(LPDISPATCH previousHotItem, LPDISPATCH newHotItem);
	void __stdcall HotItemChangingExlvwa(LPDISPATCH previousHotItem, LPDISPATCH newHotItem, VARIANT_BOOL* cancelChange);
	void __stdcall IncrementalSearchingExlvwa(BSTR currentSearchString, LONG* itemToSelect);
	void __stdcall IncrementalSearchStringChangingExlvwa(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange);
	void __stdcall InsertedColumnExlvwa(LPDISPATCH column);
	void __stdcall InsertedGroupExlvwa(LPDISPATCH Group);
	void __stdcall InsertedItemExlvwa(LPDISPATCH listItem);
	void __stdcall InsertingColumnExlvwa(LPDISPATCH column, VARIANT_BOOL* cancelInsertion);
	void __stdcall InsertingGroupExlvwa(LPDISPATCH Group, VARIANT_BOOL* cancelInsertion);
	void __stdcall InsertingItemExlvwa(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion);
	void __stdcall InvokeVerbFromSubItemControlExlvwa(LPDISPATCH listItem, BSTR verb);
	void __stdcall ItemActivateExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short shift, long x, long y);
	void __stdcall ItemAsynchronousDrawFailedExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, ExLVwLibA::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibA::ImageDrawingFailureReasonConstants failureReason, ExLVwLibA::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw);
	void __stdcall ItemBeginDragExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemBeginRDragExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemGetDisplayInfoExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* Indent, long* groupID, LPSAFEARRAY* TileViewColumns, long maxItemTextLength, BSTR* itemText, long* OverlayIndex, long* StateImageIndex, long* itemStates, VARIANT_BOOL* dontAskAgain);
	void __stdcall ItemGetGroupExlvwa(LONG itemIndex, LONG occurrenceIndex, LONG* groupIndex);
	void __stdcall ItemGetInfoTipTextExlvwa(LPDISPATCH listItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip);
	void __stdcall ItemGetOccurrencesCountExlvwa(LONG itemIndex, LONG* occurrencesCount);
	void __stdcall ItemMouseEnterExlvwa(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemMouseLeaveExlvwa(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ItemSelectionChangedExlvwa(LPDISPATCH listItem);
	void __stdcall ItemSetTextExlvwa(LPDISPATCH listItem, BSTR itemText);
	void __stdcall ItemStateImageChangedExlvwa(LPDISPATCH listItem, long previousStateImageIndex, long newStateImageIndex, long causedBy);
	void __stdcall ItemStateImageChangingExlvwa(LPDISPATCH listItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange);
	void __stdcall KeyDownExlvwa(short* keyCode, short shift);
	void __stdcall KeyPressExlvwa(short* keyAscii);
	void __stdcall KeyUpExlvwa(short* keyCode, short shift);
	void __stdcall MapGroupWideToTotalItemIndexExlvwa(LONG groupIndex, LONG groupWideItemIndex, LONG* totalItemIndex);
	void __stdcall MClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseWheelExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragExlvwa(LPDISPATCH data, long performedEffect);
	void __stdcall OLEDragDropExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragEnterExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetExlvwa(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveExlvwa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragLeavePotentialTargetExlvwa(void);
	void __stdcall OLEDragMouseMoveExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackExlvwa(long effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragExlvwa(BOOL pressedEscape, short button, short shift, long* actionToContinueWith);
	void __stdcall OLEReceivedNewDataExlvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataExlvwa(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect);
	void __stdcall OLEStartDragExlvwa(LPDISPATCH data);
	void __stdcall OwnerDrawItemExlvwa(LPDISPATCH listItem, ExLVwLibA::OwnerDrawItemStateConstants itemState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle);
	void __stdcall RClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowExlvwa(long hWnd);
	void __stdcall RemovedColumnExlvwa(LPDISPATCH column);
	void __stdcall RemovedGroupExlvwa(LPDISPATCH Group);
	void __stdcall RemovedItemExlvwa(LPDISPATCH listItem);
	void __stdcall RemovingColumnExlvwa(LPDISPATCH column, VARIANT_BOOL* cancelDeletion);
	void __stdcall RemovingGroupExlvwa(LPDISPATCH Group, VARIANT_BOOL* cancelDeletion);
	void __stdcall RemovingItemExlvwa(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion);
	void __stdcall RenamedItemExlvwa(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText);
	void __stdcall RenamingItemExlvwa(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming);
	void __stdcall ResizedControlWindowExlvwa();
	void __stdcall ResizingColumnExlvwa(LPDISPATCH column, long* newColumnWidth, VARIANT_BOOL* abortResizing);
	void __stdcall SelectedItemRangeExlvwa(LPDISPATCH firstItem, LPDISPATCH lastItem);
	void __stdcall SettingItemInfoTipTextExlvwa(LPDISPATCH listItem, BSTR* infoTipText, VARIANT_BOOL* abortInfoTip);
	void __stdcall StartingLabelEditingExlvwa(LPDISPATCH listItem, VARIANT_BOOL* cancelEditing);
	void __stdcall SubItemMouseEnterExlvwa(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall SubItemMouseLeaveExlvwa(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails);
};
