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
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	CComQIPtr<IEnumVARIANT> pEnum = pItems;
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ULONG ul = 0;
	_variant_t v;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<IListViewItem> pItem = v.pdispVal;
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
	}

	if(!IsComctl32Version600OrNewer()) {
		DeleteObject(LongToHandle(hBMPDownArrow));
		DeleteObject(LongToHandle(hBMPUpArrow));
	}

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

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.groupsCheck = GetDlgItem(IDC_GROUPSCHECK);

	exlvwUWnd.SubclassWindow(GetDlgItem(IDC_EXLVWU));
	exlvwUWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwU));
	if(controls.exlvwU) {
		DispEventAdvise(controls.exlvwU);
		InsertColumns();

		if(!IsComctl32Version600OrNewer()) {
			controls.groupsCheck.EnableWindow(FALSE);

			CComPtr<IListViewColumn> pColumn = controls.exlvwU->GetColumns()->GetItem(0, citIndex);
			pColumn->PutBitmapHandle(static_cast<OLE_HANDLE>(-1));
			hBMPDownArrow = pColumn->GetBitmapHandle();
			pColumn->PutBitmapHandle(static_cast<OLE_HANDLE>(-2));
			hBMPUpArrow = pColumn->GetBitmapHandle();
			pColumn->PutBitmapHandle(0);
		}

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

LRESULT CMainDlg::OnBnClickedGroupscheck(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutShowGroups((controls.groupsCheck.GetCheck() == BST_CHECKED) ? VARIANT_TRUE : VARIANT_FALSE);
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

int CALLBACK CMainDlg::EnumFontFamExProc(const LPENUMLOGFONTEX lpElfe, const NEWTEXTMETRICEX* /*lpntme*/, DWORD /*FontType*/, LPARAM lParam)
{
	static CAtlString lastFontName = TEXT("");

	if(lastFontName != CAtlString(lpElfe->elfFullName)) {
		LPENUMFONTPARAM pParam = reinterpret_cast<LPENUMFONTPARAM>(lParam);
		lpElfe->elfLogFont.lfWidth = pParam->lfListView.lfWidth;
		lpElfe->elfLogFont.lfHeight = pParam->lfListView.lfHeight;

		HFONT hFont = CreateFontIndirect(&lpElfe->elfLogFont);
		CComPtr<IListViewItem> pItem = pParam->pItemsToAddTo->Add(_bstr_t(lpElfe->elfFullName), -1, 0, 0, HandleToLong(hFont), pParam->lfListView.lfCharSet);
		if(pItem) {
			pItem->GetSubItems()->GetItem(1, citIndex)->PutText(_bstr_t(pParam->groupName));
		}
	}
	lastFontName = lpElfe->elfFullName;

	return 1;
}

void CMainDlg::InsertColumns(void)
{
	CComPtr<IListViewColumns> pColumns = controls.exlvwU->GetColumns();
	pColumns->Add(OLESTR("Font name"), -1, 270, 0, alLeft, 0, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE)->PutImageOnRight(VARIANT_TRUE);
	pColumns->Add(OLESTR("Charset"), -1, 150, 0, alLeft, 0, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE)->PutImageOnRight(VARIANT_TRUE);
}

void CMainDlg::InsertItems(void)
{
	// setup tooltip
	CWindow wnd = static_cast<HWND>(LongToHandle(controls.exlvwU->GethWndToolTip()));
	wnd.ModifyStyle(WS_BORDER, TTS_BALLOON);
	CToolTipCtrl tooltip = wnd;
	tooltip.SetTitle(1, TEXT("Font name:"));

	// insert groups
	if(IsComctl32Version600OrNewer()) {
		CComPtr<IListViewGroups> pGroups = controls.exlvwU->GetGroups();
		pGroups->Add(OLESTR("ANSI_CHARSET"), ANSI_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("BALTIC_CHARSET"), BALTIC_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("CHINESEBIG5_CHARSET"), CHINESEBIG5_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("DEFAULT_CHARSET"), DEFAULT_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("EASTEUROPE_CHARSET"), EASTEUROPE_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("GB2312_CHARSET"), GB2312_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("GREEK_CHARSET"), GREEK_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("HANGUL_CHARSET"), HANGUL_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("MAC_CHARSET"), MAC_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("OEM_CHARSET"), OEM_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("RUSSIAN_CHARSET"), RUSSIAN_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("SHIFTJIS_CHARSET"), SHIFTJIS_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("SYMBOL_CHARSET"), SYMBOL_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("TURKISH_CHARSET"), TURKISH_CHARSET, -1, 1, alLeft, -1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	}

	// insert items
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	if(pItems) {
		hDefaultFont = reinterpret_cast<HFONT>(SendMessage(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), WM_GETFONT, 0, 0));
		ENUMFONTPARAM param = {0};
		ZeroMemory(&param.lfListView, sizeof(LOGFONT));
		GetObject(hDefaultFont, sizeof(LOGFONT), &param.lfListView);
		param.lfListView.lfFaceName[0] = '\0';
		HDC hDC = ::GetDC(NULL);

		param.pItemsToAddTo = controls.exlvwU->GetListItems();

		param.groupName = TEXT("ANSI_CHARSET");
		param.lfListView.lfCharSet = ANSI_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("BALTIC_CHARSET");
		param.lfListView.lfCharSet = BALTIC_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("CHINESEBIG5_CHARSET");
		param.lfListView.lfCharSet = CHINESEBIG5_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("DEFAULT_CHARSET");
		param.lfListView.lfCharSet = DEFAULT_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("EASTEUROPE_CHARSET");
		param.lfListView.lfCharSet = EASTEUROPE_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("GB2312_CHARSET");
		param.lfListView.lfCharSet = GB2312_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("GREEK_CHARSET");
		param.lfListView.lfCharSet = GREEK_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("HANGUL_CHARSET");
		param.lfListView.lfCharSet = HANGUL_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("MAC_CHARSET");
		param.lfListView.lfCharSet = MAC_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("OEM_CHARSET");
		param.lfListView.lfCharSet = OEM_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("RUSSIAN_CHARSET");
		param.lfListView.lfCharSet = RUSSIAN_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("SHIFTJIS_CHARSET");
		param.lfListView.lfCharSet = SHIFTJIS_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("SYMBOL_CHARSET");
		param.lfListView.lfCharSet = SYMBOL_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		param.groupName = TEXT("TURKISH_CHARSET");
		param.lfListView.lfCharSet = TURKISH_CHARSET;
		EnumFontFamiliesEx(hDC, &param.lfListView, reinterpret_cast<FONTENUMPROC>(EnumFontFamExProc), reinterpret_cast<LPARAM>(&param), 0);

		::ReleaseDC(NULL, hDC);
	}
}

void __stdcall CMainDlg::ColumnClickExlvwu(LPDISPATCH column, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	CComQIPtr<IListViewColumn> pColumn = column;
	CComPtr<IListViewColumns> pColumns = controls.exlvwU->GetColumns();

	CComQIPtr<IEnumVARIANT> pIterator = pColumns;
	_variant_t vColumn;
	while(pIterator->Next(1, &vColumn, NULL) == S_OK) {
		if(vColumn.vt == VT_DISPATCH) {
			SortArrowConstants curArrow = saNone;
			SortArrowConstants newArrow = saNone;

			CComQIPtr<IListViewColumn> pCol = vColumn.pdispVal;
			if(pCol->GetIndex() == pColumn->GetIndex()) {
				if(IsComctl32Version600OrNewer()) {
					curArrow = pCol->GetSortArrow();
				} else {
					if(pCol->GetBitmapHandle() == 0) {
						curArrow = saNone;
					} else if(pCol->GetBitmapHandle() == static_cast<OLE_HANDLE>(hBMPDownArrow)) {
						curArrow = saDown;
					} else if(pCol->GetBitmapHandle() == static_cast<OLE_HANDLE>(hBMPUpArrow)) {
						curArrow = saUp;
					}
				}

				if(curArrow == saUp) {
					newArrow = saDown;
				} else {
					newArrow = saUp;
				}
			} else {
				newArrow = saNone;
			}

			if(IsComctl32Version600OrNewer()) {
				pCol->PutSortArrow(newArrow);
				if(newArrow == saDown) {
					controls.exlvwU->PutSortOrder(soDescending);
					controls.exlvwU->SortItems(sobShell, sobText, sobNone, sobNone, sobNone, vColumn, VARIANT_TRUE);
				} else if(newArrow == saUp) {
					controls.exlvwU->PutSortOrder(soAscending);
					controls.exlvwU->SortItems(sobShell, sobText, sobNone, sobNone, sobNone, vColumn, VARIANT_TRUE);
				}
			} else {
				if(newArrow == saNone) {
					pCol->PutBitmapHandle(0);
				} else if(newArrow == saDown) {
					pCol->PutBitmapHandle(static_cast<OLE_HANDLE>(hBMPDownArrow));
					controls.exlvwU->PutSortOrder(soDescending);
					controls.exlvwU->SortItems(sobShell, sobText, sobNone, sobNone, sobNone, vColumn, VARIANT_TRUE);
				} else if(newArrow == saUp) {
					pCol->PutBitmapHandle(static_cast<OLE_HANDLE>(hBMPUpArrow));
					controls.exlvwU->PutSortOrder(soAscending);
					controls.exlvwU->SortItems(sobShell, sobText, sobNone, sobNone, sobNone, vColumn, VARIANT_TRUE);
				}
			}
		}
		vColumn.Clear();
	}
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL /*drawAllItems*/, OLE_COLOR* /*textColor*/, OLE_COLOR* textBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants /*itemState*/, long hDC, RECTANGLE* /*drawingRectangle*/, CustomDrawReturnValuesConstants* furtherProcessing)
{
	switch(drawStage) {
		case cdsPrePaint:
			*furtherProcessing = cdrvNotifyItemDraw;
			break;

		case cdsItemPrePaint: {
			CComQIPtr<IListViewItem> pItem = listItem;
			if(pItem) {
				if(pItem->GetIndex() % 2 == 0) {
					*textBackColor = RGB(240, 240, 240);
				} else {
					*textBackColor = RGB(255, 255, 255);
				}
				*furtherProcessing = static_cast<CustomDrawReturnValuesConstants>(cdrvNewFont | cdrvNotifySubItemDraw);
			}
			break;
		}

		case cdsSubItemPrePaint: {
			CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
			if(pSubItem) {
				if(pSubItem->GetIndex() > 0) {
					SelectObject(static_cast<HDC>(LongToHandle(hDC)), static_cast<HGDIOBJ>(LongToHandle(hDefaultFont)));
					*furtherProcessing = cdrvNewFont;
				} else {
					CComQIPtr<IListViewItem> pItem = listItem;
					SelectObject(static_cast<HDC>(LongToHandle(hDC)), static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData())));
					*furtherProcessing = cdrvNewFont;
				}
			}
			break;
		}
	}
}

void __stdcall CMainDlg::FreeItemDataExlvwu(LPDISPATCH listItem)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
		if(GetObjectType(h) == OBJ_FONT) {
			DeleteObject(h);
		}
	}
}

void __stdcall CMainDlg::ItemGetInfoTipTextExlvwu(LPDISPATCH listItem, long /*maxInfoTipLength*/, BSTR* infoTipText, VARIANT_BOOL* /*abortToolTip*/)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(pItem) {
		HGDIOBJ h = static_cast<HGDIOBJ>(LongToHandle(pItem->GetItemData()));
		if(GetObjectType(h) == OBJ_FONT) {
			*infoTipText = _bstr_t(pItem->GetText()).Detach();
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
}
