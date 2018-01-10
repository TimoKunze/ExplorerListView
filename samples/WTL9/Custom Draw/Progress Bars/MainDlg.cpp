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

	KillTimer(10);

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

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	exlvwUWnd.SubclassWindow(GetDlgItem(IDC_EXLVWU));
	exlvwUWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwU));
	if(controls.exlvwU) {
		DispEventAdvise(controls.exlvwU);

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

	SetTimer(10, 250);

	return TRUE;
}

LRESULT CMainDlg::OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	for(int i = 0; i < 5; i++) {
		percentage[i] += 0.005;
		if(percentage[i] > 1.0) {
			percentage[i] = 0.0;
		}
	}
	controls.exlvwU->RedrawItems(0, 4);
	return 0;
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
	CComPtr<IListViewColumns> pColumns = controls.exlvwU->GetColumns();
	pColumns->Add(OLESTR("Items"), -1, 100, 0, alLeft, 0, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Progress"), -1, 250, 0, alRight, 0, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::InsertItems(void)
{
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	pItems->Add(OLESTR("Item 1"), -1, -2, 0, 0, -2);
	pItems->Add(OLESTR("Item 2"), -1, -2, 0, 0, -2);
	pItems->Add(OLESTR("Item 3"), -1, -2, 0, 0, -2);
	pItems->Add(OLESTR("Item 4"), -1, -2, 0, 0, -2);
	pItems->Add(OLESTR("Item 5"), -1, -2, 0, 0, -2);
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL /*drawAllItems*/, OLE_COLOR* /*textColor*/, OLE_COLOR* /*textBackColor*/, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants /*itemState*/, long hDC, RECTANGLE* drawingRectangle, CustomDrawReturnValuesConstants* furtherProcessing)
{
	switch(drawStage) {
		case cdsPrePaint:
			*furtherProcessing = cdrvNotifyItemDraw;
			break;

		case cdsItemPrePaint:
			*furtherProcessing = static_cast<CustomDrawReturnValuesConstants>(cdrvDoDefault | cdrvNotifySubItemDraw);
			break;

		case cdsSubItemPrePaint:
			*furtherProcessing = static_cast<CustomDrawReturnValuesConstants>(cdrvDoDefault | cdrvNotifyPostPaint);
			break;

		case cdsSubItemPostPaint: {
			CComQIPtr<IListViewItem> pItem = listItem;
			CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
			if(pItem && pSubItem) {
				if(pSubItem->GetIndex() == 1) {
					RECT rc = {0};
					pSubItem->GetRectangle(irtEntireItem, NULL, &rc.top, NULL, &rc.bottom);
					rc.left = drawingRectangle->Left;
					rc.top++;
					rc.bottom -= 2;
					rc.right = drawingRectangle->Right - 55;
					if(rc.right > rc.left) {
						rc.right = rc.left + static_cast<int>(percentage[pItem->GetIndex()] * static_cast<double>(rc.right - rc.left));

						FillRect(static_cast<HDC>(LongToHandle(hDC)), &rc, GetSysColorBrush(COLOR_HIGHLIGHT));
					}
				} else {
					*furtherProcessing = cdrvDoDefault;
				}
			}
			break;
		}
	}
}

void __stdcall CMainDlg::ItemGetDisplayInfoExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* /*IconIndex*/, long* /*Indent*/, long* /*groupID*/, LPSAFEARRAY* /*TileViewColumns*/, long /*maxItemTextLength*/, BSTR* itemText, long* /*OverlayIndex*/, long* /*StateImageIndex*/, long* /*itemStates*/, VARIANT_BOOL* /*dontAskAgain*/)
{
	if(requestedInfo & riItemText) {
		CComQIPtr<IListViewItem> pItem = listItem;
		CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
		if(pItem && pSubItem) {
			CAtlString str;
			str.Format(TEXT("%.1f %%"), percentage[pItem->GetIndex()] * 100);
			*itemText = _bstr_t(str).Detach();
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
}
