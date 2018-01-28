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

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	HMODULE hMod = LoadLibrary(TEXT("shell32.dll"));
	if(hMod) {
		if(IsComctl32Version600OrNewer()) {
			controls.smallImageList.Create(16, 16, ILC_COLOR32 | ILC_MASK, 10, 0);
			controls.smallImageList.SetBkColor(CLR_NONE);

			int resIDs16x16[10] = {1489, 516, 26, 185, 266, 51, 275, 126, 37, 1029};
			if(RunTimeHelper::IsVista()) {
				resIDs16x16[0] = 60;
				resIDs16x16[1] = 75;
				resIDs16x16[2] = 117;
				resIDs16x16[3] = 133;
				resIDs16x16[4] = 180;
				resIDs16x16[5] = 189;
				resIDs16x16[6] = 258;
				resIDs16x16[7] = 320;
				resIDs16x16[9] = 384;
			}

			for(int iml = 0; iml < 1; iml++) {
				int* resIDs = NULL;
				switch(iml) {
					case 0:
						resIDs = resIDs16x16;
						break;
				}

				for(int i = 0; i < 10; i++) {
					HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
					if(hResource) {
						HGLOBAL hMem = LoadResource(hMod, hResource);
						if(hMem) {
							LPVOID pIconData = LockResource(hMem);
							if(pIconData) {
								hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
								if(hIcon) {
									switch(iml) {
										case 0:
											controls.smallImageList.AddIcon(hIcon);
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

			int resIDs16x16[10] = {1486, 513, 22, 181, 263, 48, 272, 122, 34, 1026};
			if(RunTimeHelper::IsVista()) {
				resIDs16x16[0] = 57;
				resIDs16x16[1] = 72;
				resIDs16x16[2] = 114;
				resIDs16x16[3] = 129;
				resIDs16x16[4] = 177;
				resIDs16x16[5] = 186;
				resIDs16x16[6] = 255;
				resIDs16x16[7] = 317;
				resIDs16x16[9] = 381;
			}

			for(int iml = 0; iml < 1; iml++) {
				int* resIDs = NULL;
				switch(iml) {
					case 0:
						resIDs = resIDs16x16;
						break;
				}

				for(int i = 0; i < 10; i++) {
					HRSRC hResource = FindResource(hMod, MAKEINTRESOURCE(resIDs[i]), RT_ICON);
					if(hResource) {
						HGLOBAL hMem = LoadResource(hMod, hResource);
						if(hMem) {
							LPVOID pIconData = LockResource(hMem);
							if(pIconData) {
								hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
								if(hIcon) {
									switch(iml) {
										case 0:
											controls.smallImageList.AddIcon(hIcon);
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

LRESULT CMainDlg::OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutView(vList);
	controls.exlvwU->PutFullRowSelect(frsDisabled);
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

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertColumns(void)
{
	CComPtr<IListViewColumns> pColumns = controls.exlvwU->GetColumns();
	pColumns->Add(OLESTR("Column 1"), -1, 100, 0, alLeft, 1, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 2"), -1, 160, 0, alCenter, 2, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::InsertItems(void)
{
	controls.exlvwU->PuthImageList(ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
	controls.exlvwU->PutVirtualItemCount(VARIANT_FALSE, VARIANT_TRUE, 2000);
}

void CMainDlg::UpdateMenu(void)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	switch(controls.exlvwU->GetView()) {
		case vDetails:
			CheckMenuRadioItem(hMenu, ID_VIEW_LIST, ID_VIEW_DETAILS, ID_VIEW_DETAILS, MF_BYCOMMAND);
			break;
		case vList:
			CheckMenuRadioItem(hMenu, ID_VIEW_LIST, ID_VIEW_DETAILS, ID_VIEW_LIST, MF_BYCOMMAND);
			break;
	}
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::ItemGetDisplayInfoExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* /*Indent*/, long* /*groupID*/, LPSAFEARRAY* TileViewColumns, long /*maxItemTextLength*/, BSTR* itemText, long* /*OverlayIndex*/, long* StateImageIndex, long* /*itemStates*/, VARIANT_BOOL* /*dontAskAgain*/)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	int itemIndex = pItem->GetIndex();
	if(requestedInfo & riItemText) {
		CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
		if(pSubItem) {
			CAtlString tmp;
			tmp.Format(TEXT("Item %i, SubItem %i"), itemIndex, pSubItem->GetIndex());

			*itemText = _bstr_t(tmp).Detach();
		} else {
			CAtlString tmp;
			tmp.Format(TEXT("Item %i"), itemIndex);

			*itemText = _bstr_t(tmp).Detach();
		}
	}
	if(requestedInfo & riIconIndex) {
		*IconIndex = itemIndex % 10;
	}
	if(requestedInfo & riTileViewColumns) {
		CComPtr<IRecordInfo> pRecordInfo = NULL;
		CLSID clsidTILEVIEWSUBITEM = {0};
		CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
		ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, 1, 7, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
		*TileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);

		TILEVIEWSUBITEM element = {0};
		for(LONG i = 0; i < 2; ++i) {
			element.ColumnIndex = i + 1;
			element.BeginNewColumn = (i == 1 ? VARIANT_TRUE : VARIANT_FALSE);
			element.FillRemainder = VARIANT_FALSE;
			element.WrapToMultipleLines = VARIANT_TRUE;
			element.NoTitle = VARIANT_FALSE;
			ATLVERIFY(SUCCEEDED(SafeArrayPutElement(*TileViewColumns, &i, &element)));
		}
	}
	if(requestedInfo & riStateImageIndex) {
		*StateImageIndex = stateImageIndices[itemIndex] + 1;
	}
}

void __stdcall CMainDlg::ItemStateImageChangedExlvwu(LPDISPATCH listItem, long /*previousStateImageIndex*/, long newStateImageIndex, long /*causedBy*/)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	stateImageIndices[pItem->GetIndex()] = newStateImageIndex - 1;
	pItem->Update();
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
}
