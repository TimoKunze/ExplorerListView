// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::IsComctl32Version600OrNewer(void)
{
	BOOL ret = FALSE;

	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		DLLGETVERSIONPROC pDllGetVersion = reinterpret_cast<DLLGETVERSIONPROC>(GetProcAddress(hMod, "DllGetVersion"));
		if(pDllGetVersion) {
			DLLVERSIONINFO versionInfo = {0};
			versionInfo.cbSize = sizeof(versionInfo);
			HRESULT hr = (*pDllGetVersion)(&versionInfo);
			if(SUCCEEDED(hr)) {
				ret = (versionInfo.dwMajorVersion >= 6);
			}
		}
		FreeLibrary(hMod);
	}

	return ret;
}


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	if((pMsg->message == WM_KEYDOWN) && ((pMsg->wParam == VK_ESCAPE) || (pMsg->wParam == VK_RETURN))) {
		// otherwise label-edit wouldn't work
		return FALSE;
	} else {
		return CWindow::IsDialogMessage(pMsg);
	}
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.exlvwU) {
		DispEventUnadvise(controls.exlvwU);
		controls.exlvwU.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	controls.imageList.CreateFromImage(IDB_ICONS, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	exlvwUWnd.SubclassWindow(GetDlgItem(IDC_EXLVWU));
	exlvwUWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwU));
	if(controls.exlvwU) {
		DispEventAdvise(controls.exlvwU);

		// create a bold font
		hFontBold = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), WM_GETFONT, 0, 0));
		LOGFONT lf = {0};
		GetObject(hFontBold, sizeof(LOGFONT), &lf);
		lf.lfWeight |= FW_BOLD;
		hFontBold = CreateFontIndirect(&lf);

		InsertColumns();
		InsertItems();
	}

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), L"explorer", NULL);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.exlvwU) {
		controls.exlvwU->About();
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertColumns(void)
{
	RECT rc = {0};
	::GetClientRect(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), &rc);
	controls.exlvwU->GetColumns()->Add(OLESTR("Column 1"), -1, rc.right - rc.left - 10, 0, alLeft, 0, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::InsertItems(void)
{
	controls.exlvwU->PuthImageList(ilSmall, HandleToLong(controls.imageList.m_hImageList));

	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	linksID = pItems->Add(OLESTR("Links"), -1, 1, 0, 100, -2)->GetID();
	InsertLinks();
	moreLinksID = pItems->Add(OLESTR("More Links"), -1, 1, 0, 100, -2)->GetID();
	InsertMoreLinks();
}

void CMainDlg::InsertLinks(void)
{
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	long baseIndex = pItems->GetItem(linksID, 0, iitID)->GetIndex();
	pItems->Add(OLESTR("Tasks"), baseIndex + 1, 2, 2, linksID, -2);
	pItems->Add(OLESTR("Patterns"), baseIndex + 2, 3, 2, linksID, -2);
	pItems->Add(OLESTR("Deleted Items"), baseIndex + 3, 4, 2, linksID, -2);
	pItems->Add(OLESTR("Sent Items"), baseIndex + 4, 5, 2, linksID, -2);
}

void CMainDlg::InsertMoreLinks(void)
{
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	long baseIndex = pItems->GetItem(moreLinksID, 0, iitID)->GetIndex();
	pItems->Add(OLESTR("Journal"), baseIndex + 1, 6, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Junk"), baseIndex + 2, 7, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Calendar"), baseIndex + 3, 8, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Contacts"), baseIndex + 4, 9, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Notes"), baseIndex + 5, 10, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Outgoing"), baseIndex + 6, 11, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Incoming"), baseIndex + 7, 12, 2, moreLinksID, -2);
	pItems->Add(OLESTR("Search"), baseIndex + 8, 13, 2, moreLinksID, -2);
}

void CMainDlg::RemoveLinks(void)
{
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	pItems->PutFilterType(fpItemData, ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(linksID));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
	pItems->PutFilter(fpItemData, filter);
	pItems->RemoveAll();
}

void CMainDlg::RemoveMoreLinks(void)
{
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	pItems->PutFilterType(fpItemData, ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(moreLinksID));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
	pItems->PutFilter(fpItemData, filter);
	pItems->RemoveAll();
}

void __stdcall CMainDlg::ClickExlvwu(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(pItem) {
		if(hitTestDetails & htItemIcon) {
			switch(pItem->GetIconIndex()) {
				case 0:
					pItem->PutIconIndex(1);
					if(pItem->GetID() == linksID) {
						InsertLinks();
					} else if(pItem->GetID() == moreLinksID) {
						InsertMoreLinks();
					}
					break;
				case 1:
					pItem->PutIconIndex(0);
					if(pItem->GetID() == linksID) {
						RemoveLinks();
					} else if(pItem->GetID() == moreLinksID) {
						RemoveMoreLinks();
					}
					break;
			}
		}
	}
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, VARIANT_BOOL /*drawAllItems*/, OLE_COLOR* textColor, OLE_COLOR* textBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants /*itemState*/, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing)
{
	switch(drawStage) {
		case cdsPrePaint:
			*furtherProcessing = cdrvNotifyItemDraw;
			break;

		case cdsItemPrePaint: {
			*textBackColor = RGB(255, 255, 255);
			CComQIPtr<IListViewItem> pItem = listItem;
			if(pItem) {
				if(pItem->GetIndent() == 0) {
					if(hFontBold == NULL) {
						hFontBold = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), WM_GETFONT, 0, 0));
						LOGFONT lf = {0};
						GetObject(hFontBold, sizeof(LOGFONT), &lf);
						lf.lfWeight |= FW_BOLD;
						hFontBold = CreateFontIndirect(&lf);
					}
					SelectObject(static_cast<HDC>(LongToHandle(hDC)), hFontBold);

					if((pItem->GetSelected() == VARIANT_FALSE) || (GetFocus() != static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())))) {
						*textColor = GetSysColor(COLOR_HIGHLIGHT);
					}

					*furtherProcessing = cdrvNewFont;
				} else {
					*furtherProcessing = cdrvDoDefault;
				}
			}
			break;
		}
	}
}

void __stdcall CMainDlg::DblClickExlvwu(LPDISPATCH listItem, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(pItem) {
		if(hitTestDetails & htItemIcon) {
			switch(pItem->GetIconIndex()) {
				case 0:
					pItem->PutIconIndex(1);
					if(pItem->GetID() == linksID) {
						InsertLinks();
					} else if(pItem->GetID() == moreLinksID) {
						InsertMoreLinks();
					}
					break;
				case 1:
					pItem->PutIconIndex(0);
					if(pItem->GetID() == linksID) {
						RemoveLinks();
					} else if(pItem->GetID() == moreLinksID) {
						RemoveMoreLinks();
					}
					break;
			}
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
}

void __stdcall CMainDlg::StartingLabelEditingExlvwu(LPDISPATCH listItem, VARIANT_BOOL* /*cancelEditing*/)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(pItem) {
		if(pItem->GetIndent() == 0) {
			SendMessage(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWndEdit())), WM_SETFONT, reinterpret_cast<WPARAM>(hFontBold), TRUE);
			controls.exlvwU->PutEditForeColor(GetSysColor(COLOR_HIGHLIGHT));
		} else {
			controls.exlvwU->PutEditForeColor(GetSysColor(COLOR_WINDOWTEXT));
		}
	}
}
