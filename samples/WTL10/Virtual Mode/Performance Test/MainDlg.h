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
    public IDispEventImpl<IDC_EXLVWNORMAL, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>,
    public IDispEventImpl<IDC_EXLVWVIRTUAL, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> exlvwNormalContainerWnd;
	CContainedWindowT<CWindow> exlvwNormalWnd;
	CContainedWindowT<CAxWindow> exlvwVirtualContainerWnd;
	CContainedWindowT<CWindow> exlvwVirtualWnd;

	CMainDlg() :
	    exlvwNormalContainerWnd(this, 1),
	    exlvwVirtualContainerWnd(this, 2),
	    exlvwNormalWnd(this, 3),
	    exlvwVirtualWnd(this, 4)
	{
	}

	BOOL exlvwnormalIsFocused;
	long cItems;
	long cItemsHalf;

	struct Controls
	{
		CStatic descrStatic;
		CEdit countEdit;
		CButton fillListViewsButton;
		CStatic fillTimeStatic;
		CButton aboutButton;
		CImageList smallImageList;
		CImageList largeImageList;
		CImageList extralargeImageList;
		CToolTipCtrl tooltip;
		CComPtr<IExplorerListView> exlvwNormal;
		CComPtr<IExplorerListView> exlvwVirtual;

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

		COMMAND_HANDLER(IDC_FILLLISTVIEWSBUTTON, BN_CLICKED, OnBnClickedFilllistviewsbutton)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		COMMAND_ID_HANDLER(ID_VIEW_ICONS, OnViewIcons)
		COMMAND_ID_HANDLER(ID_VIEW_SMALLICONS, OnViewSmallIcons)
		COMMAND_ID_HANDLER(ID_VIEW_LIST, OnViewList)
		COMMAND_ID_HANDLER(ID_VIEW_DETAILS, OnViewDetails)
		COMMAND_ID_HANDLER(ID_VIEW_TILES, OnViewTiles)
		COMMAND_ID_HANDLER(ID_VIEW_EXTENDEDTILES, OnViewExtendedTiles)
		COMMAND_ID_HANDLER(ID_VIEW_SHOWINGROUPS, OnViewShowInGroups)
		COMMAND_ID_HANDLER(ID_ALIGN_TOP, OnAlignTop)
		COMMAND_ID_HANDLER(ID_ALIGN_LEFT, OnAlignLeft)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusNormal)

		ALT_MSG_MAP(4)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusVirtual)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXLVWNORMAL, __uuidof(_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwnormal)

		SINK_ENTRY_EX(IDC_EXLVWVIRTUAL, __uuidof(_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwvirtual)
		SINK_ENTRY_EX(IDC_EXLVWVIRTUAL, __uuidof(_IExplorerListViewEvents), 93, ItemGetDisplayInfoExlvwvirtual)
		SINK_ENTRY_EX(IDC_EXLVWVIRTUAL, __uuidof(_IExplorerListViewEvents), 166, MapGroupWideToTotalItemIndexExlvwvirtual)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnBnClickedFilllistviewsbutton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewSmallIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewDetails(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewExtendedTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnViewShowInGroups(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlignTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAlignLeft(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusNormal(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnSetFocusVirtual(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);

	void CloseDialog(int nVal);
	void InsertColumnsNormal(void);
	void InsertColumnsVirtual(void);
	void UpdateMenu(void);

	void __stdcall CreatedHeaderControlWindowExlvwnormal(long /*hWndHeader*/);

	void __stdcall CreatedHeaderControlWindowExlvwvirtual(long /*hWndHeader*/);
	void __stdcall ItemGetDisplayInfoExlvwvirtual(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* /*Indent*/, long* /*groupID*/, LPSAFEARRAY* TileViewColumns, long /*maxItemTextLength*/, BSTR* itemText, long* /*OverlayIndex*/, long* /*StateImageIndex*/, long* /*itemStates*/, VARIANT_BOOL* /*dontAskAgain*/);
	void __stdcall MapGroupWideToTotalItemIndexExlvwvirtual(LONG groupIndex, LONG groupWideItemIndex, LONG* totalItemIndex);
};
