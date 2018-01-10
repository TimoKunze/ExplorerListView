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

	CMainDlg() :
	    exlvwUWnd(this, 1)
	{
	}

	struct Controls
	{
		CButton inPlaceEditCheck;
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

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 174, ConfigureSubItemControlExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 175, EndSubItemEditExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 176, GetSubItemControlExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 177, InvokeVerbFromSubItemControlExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 124, RecreatedControlWindowExlvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_EXLVWU, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(IDC_INPLACEEDITCHECK, DLSZ_MOVE_X)
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

	void CloseDialog(int nVal);
	void InsertColumns(void);
	void InsertItems(void);
	void UpdateMenu(void);

	void __stdcall ConfigureSubItemControlExlvwu(LPDISPATCH listSubItem, SubItemControlKindConstants /*controlKind*/, SubItemEditModeConstants /*editMode*/, SubItemControlConstants subItemControl, BSTR* /*themeAppName*/, BSTR* /*themeIDList*/, long* /*hFont*/, OLE_COLOR* /*textColor*/, long* pPropertyDescription, long pPropertyValue);
	void __stdcall CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/);
	void __stdcall EndSubItemEditExlvwu(LPDISPATCH listSubItem, SubItemEditModeConstants /*editMode*/, long /*pPropertyKey*/, long pPropertyValue, VARIANT_BOOL* /*cancel*/);
	void __stdcall GetSubItemControlExlvwu(LPDISPATCH listSubItem, SubItemControlKindConstants controlKind, SubItemEditModeConstants /*editMode*/, SubItemControlConstants* subItemControl);
	void __stdcall InvokeVerbFromSubItemControlExlvwu(LPDISPATCH listItem, BSTR verb);
	void __stdcall RecreatedControlWindowExlvwu(long /*hWnd*/);
};
