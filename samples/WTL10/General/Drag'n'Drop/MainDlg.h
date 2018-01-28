// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:9FC6639B-4237-4fb5-93B8-24049D39DF74> version("1.7") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_EXLVWU, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> exlvwUWnd;

	BOOL bRightDrag;
	BOOL freeFloat;
	long lastDropTarget;
	CComQIPtr<IPicture> pDraggedPicture;
	POINT ptStartDrag;
	SIZE szHotSpotOffset;

	CMainDlg() :
	    exlvwUWnd(this, 1)
	{
		CF_SHELLIDLISTOFFSET = (CLIPFORMAT) RegisterClipboardFormat(CFSTR_SHELLIDLISTOFFSET);
		CF_TARGETCLSID = (CLIPFORMAT) RegisterClipboardFormat(CFSTR_TARGETCLSID);

		freeFloat = FALSE;
		memset(&ptStartDrag, 0, sizeof(ptStartDrag));
		memset(&szHotSpotOffset, 0, sizeof(szHotSpotOffset));
	}

	struct Controls
	{
		CButton oleDDCheck;
		CImageList smallImageList;
		CImageList largeImageList;
		CImageList extralargeImageList;
		CComPtr<IExplorerListView> exlvwU;

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

	CLIPFORMAT CF_SHELLIDLISTOFFSET;
	CLIPFORMAT CF_TARGETCLSID;

	BOOL IsComctl32Version600OrNewer(void);
	BOOL IsComctl32Version610OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(ID_VIEW_ICONS, OnViewIcons)
		COMMAND_ID_HANDLER(ID_VIEW_SMALLICONS, OnViewSmallIcons)
		COMMAND_ID_HANDLER(ID_VIEW_LIST, OnViewList)
		COMMAND_ID_HANDLER(ID_VIEW_DETAILS, OnViewDetails)
		COMMAND_ID_HANDLER(ID_VIEW_TILES, OnViewTiles)
		COMMAND_ID_HANDLER(ID_VIEW_EXTENDEDTILES, OnViewExtendedTiles)
		COMMAND_ID_HANDLER(ID_VIEW_AUTOARRANGE, OnViewAutoArrange)
		COMMAND_ID_HANDLER(ID_VIEW_SNAPTOGRID, OnViewSnapToGrid)
		COMMAND_ID_HANDLER(ID_VIEW_AERODRAGIMAGES, OnViewAeroDragImages)
		COMMAND_ID_HANDLER(ID_ALIGN_TOP, OnAlignTop)
		COMMAND_ID_HANDLER(ID_ALIGN_LEFT, OnAlignLeft)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 1, AbortedDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 10, ColumnBeginDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 24, DragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 25, DropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 52, HeaderAbortedDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 58, HeaderDragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 59, HeaderDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 68, HeaderMouseUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 69, HeaderOLECompleteDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 70, HeaderOLEDragDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 71, HeaderOLEDragEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 72, HeaderOLEDragLeaveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 73, HeaderOLEDragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 76, HeaderOLESetDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 77, HeaderOLEStartDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 91, ItemBeginDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 92, ItemBeginRDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 101, KeyDownExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 111, MouseUpExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 112, OLECompleteDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 113, OLEDragDropExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 114, OLEDragEnterExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 115, OLEDragLeaveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 116, OLEDragMouseMoveExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 119, OLESetDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 120, OLEStartDragExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 124, RecreatedControlWindowExlvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_EXLVWU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_OLEDDCHECK, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewSmallIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewDetails(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewExtendedTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewAutoArrange(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewSnapToGrid(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewAeroDragImages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlignTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlignLeft(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertColumns(void);
	void InsertItems(void);
	void UpdateMenu(void);

	void __stdcall AbortedDragExlvwu();
	void __stdcall ColumnBeginDragExlvwu(LPDISPATCH column, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/, VARIANT_BOOL* doAutomaticDragDrop);
	void __stdcall CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/);
	void __stdcall DragMouseMoveExlvwu(LPDISPATCH* dropTarget, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall DropExlvwu(LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/);
	void __stdcall HeaderAbortedDragExlvwu();
	void __stdcall HeaderDragMouseMoveExlvwu(LPDISPATCH* dropTarget, short /*button*/, short /*shift*/, long x, long y, long /*xListView*/, long /*yListView*/, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall HeaderDropExlvwu(LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/);
	void __stdcall HeaderMouseUpExlvwu(LPDISPATCH /*column*/, short button, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall HeaderOLECompleteDragExlvwu(LPDISPATCH data, long /*performedEffect*/);
	void __stdcall HeaderOLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH /*dropTarget*/, short /*button*/, short shift, long x, long y, long /*xListView*/, long /*yListView*/, long /*hitTestDetails*/);
	void __stdcall HeaderOLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*xListView*/, long /*yListView*/, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall HeaderOLEDragLeaveExlvwu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*xListView*/, long /*yListView*/, long /*hitTestDetails*/);
	void __stdcall HeaderOLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*xListView*/, long /*yListView*/, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall HeaderOLESetDataExlvwu(LPDISPATCH data, long formatID, long /*Index*/, long /*dataOrViewAspect*/);
	void __stdcall HeaderOLEStartDragExlvwu(LPDISPATCH data);
	void __stdcall ItemBeginDragExlvwu(LPDISPATCH /*listItem*/, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/);
	void __stdcall ItemBeginRDragExlvwu(LPDISPATCH /*listItem*/, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/);
	void __stdcall KeyDownExlvwu(short* keyCode, short /*shift*/);
	void __stdcall MouseUpExlvwu(LPDISPATCH /*listItem*/, LPDISPATCH /*listSubItem*/, short button, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails);
	void __stdcall OLECompleteDragExlvwu(LPDISPATCH data, long /*performedEffect*/);
	void __stdcall OLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* /*dropTarget*/, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/);
	void __stdcall OLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall OLEDragLeaveExlvwu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall OLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/);
	void __stdcall OLESetDataExlvwu(LPDISPATCH data, long formatID, long /*Index*/, long /*dataOrViewAspect*/);
	void __stdcall OLEStartDragExlvwu(LPDISPATCH data);
	void __stdcall RecreatedControlWindowExlvwu(long /*hWnd*/);
};
