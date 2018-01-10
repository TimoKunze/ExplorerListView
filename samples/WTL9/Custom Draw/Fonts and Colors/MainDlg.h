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

	long hBMPDownArrow;
	long hBMPUpArrow;
	HFONT hDefaultFont;

	typedef struct ENUMFONTPARAM
	{
		LOGFONT lfListView;
		CComPtr<IListViewItems> pItemsToAddTo;
		CAtlString groupName;
	} ENUMFONTPARAM, *LPENUMFONTPARAM;
	static int CALLBACK EnumFontFamExProc(const LPENUMLOGFONTEX lpElfe, const NEWTEXTMETRICEX* /*lpntme*/, DWORD /*FontType*/, LPARAM lParam);

	CMainDlg() :
	    exlvwUWnd(this, 1)
	{
		hBMPDownArrow = NULL;
		hBMPUpArrow = NULL;
		hDefaultFont = NULL;
	}

	struct Controls
	{
		CButton groupsCheck;
		CComPtr<IExplorerListView> exlvwU;
	} controls;

	BOOL IsComctl32Version600OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		COMMAND_HANDLER(IDC_GROUPSCHECK, BN_CLICKED, OnBnClickedGroupscheck)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 19, CustomDrawExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 50, FreeItemDataExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 94, ItemGetInfoTipTextExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 124, RecreatedControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 142, ColumnClickExlvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedGroupscheck(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertColumns(void);
	void InsertItems(void);

	void __stdcall ColumnClickExlvwu(LPDISPATCH column, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/);
	void __stdcall CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/);
	void __stdcall CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL /*drawAllItems*/, OLE_COLOR* /*textColor*/, OLE_COLOR* textBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants /*itemState*/, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall FreeItemDataExlvwu(LPDISPATCH listItem);
	void __stdcall ItemGetInfoTipTextExlvwu(LPDISPATCH listItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/);
	void __stdcall RecreatedControlWindowExlvwu(long /*hWnd*/);
};
