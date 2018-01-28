// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"

#pragma comment(lib, "propsys.lib")


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

BOOL CMainDlg::IsComctl32Version610OrNewer(void)
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
				if(versionInfo.dwMajorVersion == 6) {
					ret = (versionInfo.dwMinorVersion >= 10);
				} else {
					ret = (versionInfo.dwMajorVersion > 6);
				}
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

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	HMODULE hMod = LoadLibrary(TEXT("shell32.dll"));
	if(hMod) {
		if(IsComctl32Version600OrNewer()) {
			controls.smallImageList.Create(16, 16, ILC_COLOR32 | ILC_MASK, 10, 0);
			controls.smallImageList.SetBkColor(CLR_NONE);
			controls.largeImageList.Create(32, 32, ILC_COLOR32 | ILC_MASK, 10, 0);
			controls.largeImageList.SetBkColor(CLR_NONE);
			controls.extralargeImageList.Create(48, 48, ILC_COLOR32 | ILC_MASK, 10, 0);
			controls.extralargeImageList.SetBkColor(CLR_NONE);

			int resIDs16x16[10] = {1489, 516, 26, 185, 266, 51, 275, 126, 37, 1029};
			int resIDs32x32[10] = {1488, 515, 24, 183, 265, 50, 274, 124, 36, 1028};
			int resIDs48x48[10] = {1487, 514, 23, 182, 264, 49, 273, 123, 35, 1027};
			if(RunTimeHelper::IsVista()) {
				resIDs16x16[0] = 60;
				resIDs32x32[0] = 59;
				resIDs48x48[0] = 58;
				resIDs16x16[1] = 75;
				resIDs32x32[1] = 74;
				resIDs48x48[1] = 73;
				resIDs16x16[2] = 117;
				resIDs32x32[2] = 116;
				resIDs48x48[2] = 115;
				resIDs16x16[3] = 133;
				resIDs32x32[3] = 132;
				resIDs48x48[3] = 131;
				resIDs16x16[4] = 180;
				resIDs32x32[4] = 179;
				resIDs48x48[4] = 178;
				resIDs16x16[5] = 189;
				resIDs32x32[5] = 188;
				resIDs48x48[5] = 187;
				resIDs16x16[6] = 258;
				resIDs32x32[6] = 257;
				resIDs48x48[6] = 256;
				resIDs16x16[7] = 320;
				resIDs32x32[7] = 319;
				resIDs48x48[7] = 318;
				resIDs16x16[9] = 384;
				resIDs32x32[9] = 383;
				resIDs48x48[9] = 382;
			}

			for(int iml = 0; iml < 3; iml++) {
				int* resIDs = NULL;
				switch(iml) {
					case 0:
						resIDs = resIDs16x16;
						break;
					case 1:
						resIDs = resIDs32x32;
						break;
					case 2:
						resIDs = resIDs48x48;
						break;
				}

				for(int i = 0; i < 10; i++) {
					HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
					if(hResource) {
						HGLOBAL hMem = LoadResource(hMod, hResource);
						if(hMem) {
							LPVOID pIconData = LockResource(hMem);
							if(pIconData) {
								hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR | LR_CREATEDIBSECTION);
								if(hIcon) {
									switch(iml) {
										case 0:
											controls.smallImageList.AddIcon(hIcon);
											break;
										case 1:
											controls.largeImageList.AddIcon(hIcon);
											break;
										case 2:
											controls.extralargeImageList.AddIcon(hIcon);
											break;
									}
									DestroyIcon(hIcon);
								}
							}
						}
					}
				}
			}
		} else {
			controls.smallImageList.Create(16, 16, ILC_COLOR24 | ILC_MASK, 10, 0);
			controls.smallImageList.SetBkColor(CLR_NONE);
			controls.largeImageList.Create(32, 32, ILC_COLOR24 | ILC_MASK, 10, 0);
			controls.largeImageList.SetBkColor(CLR_NONE);

			int resIDs16x16[10] = {1486, 513, 22, 181, 263, 48, 272, 122, 34, 1026};
			int resIDs32x32[10] = {1485, 512, 20, 179, 262, 47, 271, 120, 33, 1025};
			if(RunTimeHelper::IsVista()) {
				resIDs16x16[0] = 57;
				resIDs32x32[0] = 56;
				resIDs16x16[1] = 72;
				resIDs32x32[1] = 71;
				resIDs16x16[2] = 114;
				resIDs32x32[2] = 113;
				resIDs16x16[3] = 129;
				resIDs32x32[3] = 128;
				resIDs16x16[4] = 177;
				resIDs32x32[4] = 176;
				resIDs16x16[5] = 186;
				resIDs32x32[5] = 185;
				resIDs16x16[6] = 255;
				resIDs32x32[6] = 254;
				resIDs16x16[7] = 317;
				resIDs32x32[7] = 316;
				resIDs16x16[9] = 381;
				resIDs32x32[9] = 380;
			}

			for(int iml = 0; iml < 2; iml++) {
				int* resIDs = NULL;
				switch(iml) {
					case 0:
						resIDs = resIDs16x16;
						break;
					case 1:
						resIDs = resIDs32x32;
						break;
				}

				for(int i = 0; i < 10; i++) {
					HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
					if(hResource) {
						HGLOBAL hMem = LoadResource(hMod, hResource);
						if(hMem) {
							LPVOID pIconData = LockResource(hMem);
							if(pIconData) {
								hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR | LR_CREATEDIBSECTION);
								if(hIcon) {
									switch(iml) {
										case 0:
											controls.smallImageList.AddIcon(hIcon);
											break;
										case 1:
											controls.largeImageList.AddIcon(hIcon);
											break;
									}
									DestroyIcon(hIcon);
								}
							}
						}
					}
				}
			}
		}
		FreeLibrary(hMod);
	}

	controls.inPlaceEditCheck = GetDlgItem(IDC_INPLACEEDITCHECK);
	controls.inPlaceEditCheck.SetCheck(BST_CHECKED);
	exlvwUWnd.SubclassWindow(GetDlgItem(IDC_EXLVWU));
	exlvwUWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwU));
	if(controls.exlvwU) {
		DispEventAdvise(controls.exlvwU);

		InsertColumns();
		InsertItems();
		UpdateMenu();
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

LRESULT CMainDlg::OnViewIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vIcons);
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewSmallIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vSmallIcons);
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vList);
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewDetails(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vDetails);
	controls.exlvwU->PutFullRowSelect(frsExtendedMode);
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vTiles);
	controls.exlvwU->PutFullRowSelect(frsDisabled);
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewExtendedTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vExtendedTiles);
	controls.exlvwU->PutFullRowSelect(frsDisabled);
	UpdateMenu();
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
	pColumns->Add(OLESTR("Column 1"), -1, 160, 0, alLeft, 1, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 2"), -1, 200, 0, alLeft, 2, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 3"), -1, 160, 0, alLeft, 3, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::InsertItems(void)
{
	controls.exlvwU->PuthImageList(ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
	controls.exlvwU->PuthImageList(ilLarge, HandleToLong(controls.largeImageList.m_hImageList));
	controls.exlvwU->PuthImageList(ilExtraLarge, HandleToLong(controls.extralargeImageList.m_hImageList));

	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	CComPtr<IRecordInfo> pRecordInfo = NULL;
	CLSID clsidTILEVIEWSUBITEM = {0};
	CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
	ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, 1, 7, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
	LPSAFEARRAY pTileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);
	TILEVIEWSUBITEM element = {0};
	for(LONG i = 0; i < 2; ++i) {
		element.ColumnIndex = i + 1;
		element.BeginNewColumn = (i == 1 ? VARIANT_TRUE : VARIANT_FALSE);
		element.FillRemainder = VARIANT_FALSE;
		element.WrapToMultipleLines = VARIANT_TRUE;
		element.NoTitle = VARIANT_FALSE;
		ATLVERIFY(SUCCEEDED(SafeArrayPutElement(pTileViewColumns, &i, &element)));
	}
	_variant_t tileViewColumns12;
	tileViewColumns12.parray = pTileViewColumns;
	tileViewColumns12.vt = VT_ARRAY | VT_RECORD;

	pTileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);
	for(LONG i = 0; i < 2; ++i) {
		element.ColumnIndex = 2 - i;
		element.BeginNewColumn = (i == 1 ? VARIANT_TRUE : VARIANT_FALSE);
		element.FillRemainder = VARIANT_FALSE;
		element.WrapToMultipleLines = VARIANT_TRUE;
		element.NoTitle = VARIANT_FALSE;
		ATLVERIFY(SUCCEEDED(SafeArrayPutElement(pTileViewColumns, &i, &element)));
	}
	_variant_t tileViewColumns21;
	tileViewColumns21.parray = pTileViewColumns;
	tileViewColumns21.vt = VT_ARRAY | VT_RECORD;

	CComPtr<IListViewSubItems> pSubItems = pItems->Add(OLESTR("No Sub-Item Control"), -1, 0, 0, 0, -2, tileViewColumns12)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("Hello World"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Multi-Line Text"), -1, 1, 0, sicMultiLineText, -2, tileViewColumns21)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("This text\nconsists of\n3 lines"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Percent Bar"), -1, 2, 0, sicPercentBar, -2, tileViewColumns12)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("46"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Rating"), -1, 3, 0, sicRating, -2, tileViewColumns21)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("3"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Text"), -1, 4, 0, sicText, -2, tileViewColumns12)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("Some Text"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Boolean Check Mark"), -1, 5, 0, sicBooleanCheckMark, -2, tileViewColumns21)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("-1"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Calendar"), -1, 6, 0, sicCalendar, -2, tileViewColumns12)->GetSubItems();
	DATE today;
	SYSTEMTIME systemTime = {0};
	GetSystemTime(&systemTime);
	systemTime.wMilliseconds = systemTime.wSecond = systemTime.wMinute = systemTime.wHour = 0;
	SystemTimeToVariantTime(&systemTime, &today);
	BSTR todayAsString;
	VarBstrFromDate(today, 0, LOCALE_NOUSEROVERRIDE, &todayAsString);
	pSubItems->GetItem(1, citIndex)->PutText(todayAsString);
	SysFreeString(todayAsString);
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Boolean Drop-Down List"), -1, 7, 0, sicCheckboxDropList, -2, tileViewColumns21)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("-1"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Drop-Down List"), -1, 8, 0, sicDropList, -2, tileViewColumns12)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("2"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));

	pSubItems = pItems->Add(OLESTR("Hyperlink"), -1, 9, 0, sicHyperlink, -2, tileViewColumns21)->GetSubItems();
	pSubItems->GetItem(1, citIndex)->PutText(OLESTR("<a id=\"Open http://www.timosoft-software.de\">http://www.timosoft-software.de</a>"));
	pSubItems->GetItem(2, citIndex)->PutText(OLESTR("Normal Sub-Item"));
}

void CMainDlg::UpdateMenu(void)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	switch(controls.exlvwU->GetView()) {
		case vDetails:
			CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_DETAILS, MF_BYCOMMAND);
			break;
		case vIcons:
			CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_ICONS, MF_BYCOMMAND);
			break;
		case vList:
			CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_LIST, MF_BYCOMMAND);
			break;
		case vSmallIcons:
			CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_SMALLICONS, MF_BYCOMMAND);
			break;
		case vTiles:
			CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_TILES, MF_BYCOMMAND);
			break;
		case vExtendedTiles:
			CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND);
			break;
	}
	EnableMenuItem(hMenu, ID_VIEW_TILES, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
	EnableMenuItem(hMenu, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
}

void __stdcall CMainDlg::ConfigureSubItemControlExlvwu(LPDISPATCH listSubItem, SubItemControlKindConstants /*controlKind*/, SubItemEditModeConstants /*editMode*/, SubItemControlConstants subItemControl, BSTR* /*themeAppName*/, BSTR* /*themeIDList*/, long* /*hFont*/, OLE_COLOR* /*textColor*/, long* pPropertyDescription, long pPropertyValue)
{
	// NOTE: With some more effort, you could provide custom implementations of the IPropertyDescription interface
	switch(subItemControl) {
		case sicDropList:
			CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
			if(pSubItem) {
				IPropertyDescription** pPropDesc = reinterpret_cast<IPropertyDescription**>(pPropertyDescription);
				PSGetPropertyDescriptionByName(L"System.Photo.MeteringMode", IID_IPropertyDescription, reinterpret_cast<LPVOID*>(pPropDesc));
				PROPVARIANT* pPropVal = *reinterpret_cast<PROPVARIANT**>(&pPropertyValue);
				PROPVARIANT pv;
				PropVariantInit(&pv);
				pv.vt = VT_BSTR;
				pv.bstrVal = pSubItem->GetText();
				PropVariantChangeType(pPropVal, pv, 0, VT_I4);
				PropVariantClear(&pv);
			}
			break;
	}
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::EndSubItemEditExlvwu(LPDISPATCH listSubItem, SubItemEditModeConstants /*editMode*/, long /*pPropertyKey*/, long pPropertyValue, VARIANT_BOOL* /*cancel*/)
{
	PROPVARIANT* pPropVal = *reinterpret_cast<PROPVARIANT**>(&pPropertyValue);
	CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR str = NULL;
		PropVariantToBSTR(*pPropVal, &str);
		pSubItem->PutText(str);
		SysFreeString(str);
	}
}

void __stdcall CMainDlg::GetSubItemControlExlvwu(LPDISPATCH listSubItem, SubItemControlKindConstants controlKind, SubItemEditModeConstants /*editMode*/, SubItemControlConstants* subItemControl)
{
	CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		if(pSubItem->GetIndex() == 1) {
			if((controlKind != sickInPlaceEditing) || (controls.inPlaceEditCheck.GetCheck() == BST_CHECKED)) {
        *subItemControl = static_cast<SubItemControlConstants>(pSubItem->GetParentItem()->GetItemData());
			}
		}
	}
}

void __stdcall CMainDlg::InvokeVerbFromSubItemControlExlvwu(LPDISPATCH listItem, BSTR verb)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(pItem) {
		CAtlString str;
		BSTR tmp = pItem->GetText();
		str.Format(TEXT("TODO for item '%s': %s"), tmp, verb);
		SysFreeString(tmp);
		MessageBox(str);
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
	UpdateMenu();
}
