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
	pLoop->RemoveIdleHandler(this);
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
	controls.stateImageList.CreateFromImage(IDB_STATEIMAGES, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	exlvwUWnd.SubclassWindow(GetDlgItem(IDC_EXLVWU));
	exlvwUWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwU));
	if(controls.exlvwU) {
		DispEventAdvise(controls.exlvwU);

		/* NOTE: If we set the state image list before the control sets the LVS_EX_CHECKBOXES style, custom
		         state image won't work, so we use the OnIdle handler. */
		/*InsertColumns();
		InsertItems();*/
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

BOOL CMainDlg::OnIdle()
{
	if(!isInitialized) {
		/* NOTE: If we set the state image list before the control sets the LVS_EX_CHECKBOXES style, custom
		         state image won't work, so we use the OnIdle handler. */
		InsertColumns();
		InsertItems();
	}
	return TRUE;
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
	controls.exlvwU->PuthImageList(ilState, HandleToLong(controls.stateImageList.m_hImageList));

	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	if(pItems) {
		pItems->Add(OLESTR("Check 1"), -1, 0, 0, 1, -2)->PutStateImageIndex(1);
		pItems->Add(OLESTR("Option 1"), -1, 1, 1, 2, -2)->PutStateImageIndex(3);
		pItems->Add(OLESTR("Option 2"), -1, 2, 1, 2, -2)->PutStateImageIndex(3);
		pItems->Add(OLESTR("Option 3"), -1, 3, 1, 2, -2)->PutStateImageIndex(3);
		pItems->Add(OLESTR("Check 2"), -1, 4, 0, 1, -2)->PutStateImageIndex(1);
		pItems->Add(OLESTR("Option 1"), -1, 5, 1, 3, -2)->PutStateImageIndex(3);
		pItems->Add(OLESTR("Option 2"), -1, 6, 1, 3, -2)->PutStateImageIndex(3);
		pItems->Add(OLESTR("Option 3"), -1, 7, 1, 3, -2)->PutStateImageIndex(3);
	}
	isInitialized = TRUE;
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::ItemStateImageChangingExlvwu(LPDISPATCH listItem, long previousStateImageIndex, long* newStateImageIndex, long /*causedBy*/, VARIANT_BOOL* /*cancelChange*/)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	switch(pItem->GetItemData()) {
		case 1:
			// CheckBox
			if(previousStateImageIndex == 1) {
				// check the item
				*newStateImageIndex = 2;
			} else {
				// un-check the item
				*newStateImageIndex = 1;
			}
			break;
		case 2:
		case 3:
			// OptionButton
			if(previousStateImageIndex == 3) {
				// uncheck all items that have the same ItemData value
				CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
				pItems->PutFilterType(fpItemData, ftIncluding);
				_variant_t filter;
				filter.Clear();
				CComSafeArray<VARIANT> arr;
				arr.Create(1, 1);
				arr.SetAt(1, _variant_t(pItem->GetItemData()));
				filter.parray = arr.Detach();
				filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
				pItems->PutFilter(fpItemData, filter);

				pItems->PutFilterType(fpStateImageIndex, ftIncluding);
				filter.Clear();
				arr.Create(1, 1);
				arr.SetAt(1, _variant_t(4));
				filter.parray = arr.Detach();
				filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
				pItems->PutFilter(fpStateImageIndex, filter);

				CComQIPtr<IEnumVARIANT> pEnum = pItems;
				pEnum->Reset();     // NOTE: It's important to call Reset()!
				ULONG ul = 0;
				_variant_t v;
				v.Clear();
				while(pEnum->Next(1, &v, &ul) == S_OK) {
					pItem = v.pdispVal;
					pItem->PutStateImageIndex(3);
					v.Clear();
				}

				// check the item
				*newStateImageIndex = 4;
			} else {
				// leave the item checked
				*newStateImageIndex = 4;
			}
			break;
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
}
