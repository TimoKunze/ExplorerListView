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

	HFONT hFontBold;
	long linksID;
	long moreLinksID;

	CMainDlg() :
	    exlvwUWnd(this, 1)
	{
		linksID = -1;
		moreLinksID = -1;
	}

	struct Controls
	{
		CImageList imageList;
		CComPtr<IExplorerListView> exlvwU;

		~Controls()
		{
			if(!imageList.IsNull()) {
				imageList.Destroy();
			}
		}
	} controls;

	BOOL IsComctl32Version600OrNewer(void);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 9, ClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 18, CreatedHeaderControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 19, CustomDrawExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 20, DblClickExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 124, RecreatedControlWindowExlvwu)
		SINK_ENTRY_EX(IDC_EXLVWU, __uuidof(_IExplorerListViewEvents), 136, StartingLabelEditingExlvwu)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
	void InsertColumns(void);
	void InsertItems(void);
	void InsertLinks(void);
	void InsertMoreLinks(void);
	void RemoveLinks(void);
	void RemoveMoreLinks(void);

	void __stdcall ClickExlvwu(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails);
	void __stdcall CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/);
	void __stdcall CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, VARIANT_BOOL /*drawAllItems*/, OLE_COLOR* textColor, OLE_COLOR* textBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants /*itemState*/, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing);
	void __stdcall DblClickExlvwu(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails);
	void __stdcall RecreatedControlWindowExlvwu(long /*hWnd*/);
	void __stdcall StartingLabelEditingExlvwu(LPDISPATCH listItem, VARIANT_BOOL* /*cancelEditing*/);
};
