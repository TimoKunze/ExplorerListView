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
	if(controls.exlvwNormal) {
		IDispEventImpl<IDC_EXLVWNORMAL, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>::DispEventUnadvise(controls.exlvwNormal);
		controls.exlvwNormal.Release();
	}
	if(controls.exlvwVirtual) {
		IDispEventImpl<IDC_EXLVWVIRTUAL, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>::DispEventUnadvise(controls.exlvwVirtual);
		controls.exlvwVirtual.Release();
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
								hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
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
								hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
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

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.descrStatic = GetDlgItem(IDC_DESCRSTATIC);
	controls.countEdit = GetDlgItem(IDC_COUNTEDIT);
	controls.countEdit.SetWindowText(TEXT("1000"));
	controls.fillListViewsButton = GetDlgItem(IDC_FILLLISTVIEWSBUTTON);
	controls.fillTimeStatic = GetDlgItem(IDC_FILLTIMESTATIC);
	controls.fillTimeStatic.SetWindowText(TEXT("Fill time: 0 ms/0 ms"));
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);
	controls.tooltip.Create(*this, 0, NULL, WS_OVERLAPPED | WS_POPUP | WS_VISIBLE | TTS_NOPREFIX | TTS_BALLOON, WS_EX_LEFT | WS_EX_LTRREADING | WS_EX_TOOLWINDOW);
	controls.tooltip.Activate(TRUE);

	exlvwNormalContainerWnd.SubclassWindow(GetDlgItem(IDC_EXLVWNORMAL));
	exlvwNormalContainerWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwNormal));
	if(controls.exlvwNormal) {
		IDispEventImpl<IDC_EXLVWNORMAL, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>::DispEventAdvise(controls.exlvwNormal);
		exlvwNormalWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.exlvwNormal->GethWnd())));
		CToolInfo ti(TTF_SUBCLASS | TTF_CENTERTIP, exlvwNormalWnd, 0, NULL, TEXT("Common ListView"));
		controls.tooltip.AddTool(ti);

		controls.exlvwNormal->PuthImageList(ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
		controls.exlvwNormal->PuthImageList(ilLarge, HandleToLong(controls.largeImageList.m_hImageList));
		controls.exlvwNormal->PuthImageList(ilExtraLarge, HandleToLong(controls.extralargeImageList.m_hImageList));
		InsertColumnsNormal();
	}
	exlvwVirtualContainerWnd.SubclassWindow(GetDlgItem(IDC_EXLVWVIRTUAL));
	exlvwVirtualContainerWnd.QueryControl(__uuidof(IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwVirtual));
	if(controls.exlvwVirtual) {
		IDispEventImpl<IDC_EXLVWVIRTUAL, CMainDlg, &__uuidof(_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>::DispEventAdvise(controls.exlvwVirtual);
		exlvwVirtualWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.exlvwVirtual->GethWnd())));
		CToolInfo ti(TTF_SUBCLASS | TTF_CENTERTIP, exlvwVirtualWnd, 0, NULL, TEXT("Virtual ListView"));
		controls.tooltip.AddTool(ti);

		controls.exlvwVirtual->PuthImageList(ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
		controls.exlvwVirtual->PuthImageList(ilLarge, HandleToLong(controls.largeImageList.m_hImageList));
		controls.exlvwVirtual->PuthImageList(ilExtraLarge, HandleToLong(controls.extralargeImageList.m_hImageList));
		InsertColumnsVirtual();
	}

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.exlvwNormal->GethWnd())), L"explorer", NULL);
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.exlvwVirtual->GethWnd())), L"explorer", NULL);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	// force control resize
	WINDOWPOS dummy = {0};
	BOOL b = FALSE;
	OnWindowPosChanged(WM_WINDOWPOSCHANGED, 0, reinterpret_cast<LPARAM>(&dummy), b);

	return TRUE;
}

LRESULT CMainDlg::OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.countEdit && controls.fillListViewsButton && controls.fillTimeStatic && controls.aboutButton) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect rc;
			GetClientRect(&rc);
			int cx = static_cast<int>(0.5 * static_cast<double>(rc.Width()));
			controls.descrStatic.SetWindowPos(NULL, 3, rc.Height() - 26, 0, 0, SWP_NOSIZE);
			controls.countEdit.SetWindowPos(NULL, 40, rc.Height() - 30, 0, 0, SWP_NOSIZE);
			controls.fillListViewsButton.SetWindowPos(NULL, 95, rc.Height() - 30, 0, 0, SWP_NOSIZE);
			controls.fillTimeStatic.SetWindowPos(NULL, 200, rc.Height() - 26, 0, 0, SWP_NOSIZE);
			controls.aboutButton.SetWindowPos(NULL, 370, rc.Height() - 30, 0, 0, SWP_NOSIZE);
			exlvwVirtualContainerWnd.SetWindowPos(NULL, 0, 0, cx - 3, rc.Height() - 36, SWP_NOMOVE);
			exlvwNormalContainerWnd.SetWindowPos(NULL, cx + 3, 0, cx - 3, rc.Height() - 36, 0);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnBnClickedFilllistviewsbutton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.fillListViewsButton.EnableWindow(FALSE);

	int len = controls.countEdit.GetWindowTextLength() + 1;
	LPTSTR pBuffer = reinterpret_cast<LPTSTR>(malloc(len * sizeof(TCHAR)));
	controls.countEdit.GetWindowText(pBuffer, len);
	cItems = _ttol(pBuffer);
	cItemsHalf = cItems >> 1;
	free(pBuffer);

	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	BOOL showInGroups = ((GetMenuState(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND) & MF_CHECKED) == MF_CHECKED);

	DWORD dwStartVirtual = GetTickCount();
	controls.exlvwVirtual->PutDontRedraw(VARIANT_TRUE);
	if(showInGroups) {
		CComQIPtr<IListViewGroups> pGroups = controls.exlvwVirtual->GetGroups();
		pGroups->RemoveAll();
		pGroups->Add(OLESTR("Group 1"), 1, -1, cItemsHalf, alLeft, -1, VARIANT_TRUE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("Group 2"), 2, -1, cItems - cItemsHalf, alLeft, -1, VARIANT_TRUE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	}
	controls.exlvwVirtual->PutShowGroups(showInGroups ? VARIANT_TRUE : VARIANT_FALSE);
	controls.exlvwVirtual->PutVirtualItemCount(VARIANT_FALSE, VARIANT_TRUE, cItems);
	controls.exlvwVirtual->PutDontRedraw(VARIANT_FALSE);
	DWORD dwEndVirtual = GetTickCount();


	DWORD dwStartNormal = GetTickCount();
	controls.exlvwNormal->PutDontRedraw(VARIANT_TRUE);
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
	_variant_t tileViewColumns;
	tileViewColumns.parray = pTileViewColumns;
	tileViewColumns.vt = VT_ARRAY | VT_RECORD;
	if(showInGroups) {
		CComQIPtr<IListViewGroups> pGroups = controls.exlvwNormal->GetGroups();
		pGroups->RemoveAll();
		pGroups->Add(OLESTR("Group 1"), 1, -1, 1, alLeft, -1, VARIANT_TRUE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
		pGroups->Add(OLESTR("Group 2"), 2, -1, 1, alLeft, -1, VARIANT_TRUE, VARIANT_FALSE, OLESTR(""), alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	}
	controls.exlvwNormal->PutShowGroups(showInGroups ? VARIANT_TRUE : VARIANT_FALSE);
	CComPtr<IListViewItems> pItems = controls.exlvwNormal->GetListItems();
	pItems->RemoveAll();

	CAtlString str;
	long groupID = (showInGroups ? 1 : -2);
	for(long i = 0; i < cItems; i++) {
		if(showInGroups && i == cItemsHalf) {
			groupID = 2;
		}
		str.Format(TEXT("Item %i"), i);
		CComPtr<IListViewSubItems> pSubItems = pItems->Add(_bstr_t(str), -1, i % 10, 0, 0, groupID, tileViewColumns)->GetSubItems();
		CComPtr<IListViewSubItem> pSubItem = pSubItems->GetItem(1, citIndex);
		str.Format(TEXT("Item %i, SuItem 1"), i);
		pSubItem->PutText(_bstr_t(str));
		pSubItem->PutIconIndex((i + 1) % 10);
		pSubItem = pSubItems->GetItem(2, citIndex);
		str.Format(TEXT("Item %i, SuItem 2"), i);
		pSubItem->PutText(_bstr_t(str));
		pSubItem->PutIconIndex((i + 1) % 10);
	}
	controls.exlvwNormal->PutDontRedraw(VARIANT_FALSE);
	DWORD dwEndNormal = GetTickCount();

	str.Format(TEXT("Fill time: %i ms/%i ms"), dwEndVirtual - dwStartVirtual, dwEndNormal - dwStartNormal);
	controls.fillTimeStatic.SetWindowText(str);
	controls.fillListViewsButton.EnableWindow(TRUE);
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		if(controls.exlvwNormal) {
			controls.exlvwNormal->About();
		}
	} else {
		if(controls.exlvwVirtual) {
			controls.exlvwVirtual->About();
		}
	}
	return 0;
}

LRESULT CMainDlg::OnViewIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutView(vIcons);
	} else {
		controls.exlvwVirtual->PutView(vIcons);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewSmallIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutView(vSmallIcons);
	} else {
		controls.exlvwVirtual->PutView(vSmallIcons);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutView(vList);
	} else {
		controls.exlvwVirtual->PutView(vList);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewDetails(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutView(vDetails);
		controls.exlvwNormal->PutFullRowSelect(frsExtendedMode);
	} else {
		controls.exlvwVirtual->PutView(vDetails);
		controls.exlvwVirtual->PutFullRowSelect(frsExtendedMode);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutView(vTiles);
		controls.exlvwNormal->PutFullRowSelect(frsDisabled);
	} else {
		controls.exlvwVirtual->PutView(vTiles);
		controls.exlvwVirtual->PutFullRowSelect(frsDisabled);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewExtendedTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutView(vExtendedTiles);
		controls.exlvwNormal->PutFullRowSelect(frsDisabled);
	} else {
		controls.exlvwVirtual->PutView(vExtendedTiles);
		controls.exlvwVirtual->PutFullRowSelect(frsDisabled);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewShowInGroups(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(GetMenuState(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND) & MF_CHECKED) {
		CheckMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | MF_UNCHECKED);
	} else {
		CheckMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | MF_CHECKED);
	}
	return 0;
}

LRESULT CMainDlg::OnAlignTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutItemAlignment(iaTop);
	} else {
		controls.exlvwVirtual->PutItemAlignment(iaTop);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnAlignLeft(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwnormalIsFocused) {
		controls.exlvwNormal->PutItemAlignment(iaLeft);
	} else {
		controls.exlvwVirtual->PutItemAlignment(iaLeft);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnSetFocusNormal(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	exlvwnormalIsFocused = TRUE;
	UpdateMenu();
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusVirtual(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	exlvwnormalIsFocused = FALSE;
	UpdateMenu();
	bHandled = FALSE;
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertColumnsNormal(void)
{
	CComPtr<IListViewColumns> pColumns = controls.exlvwNormal->GetColumns();
	pColumns->Add(OLESTR("Column 1"), -1, 160, 0, alLeft, 1, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 2"), -1, 160, 0, alCenter, 2, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 3"), -1, 160, 0, alCenter, 3, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::InsertColumnsVirtual(void)
{
	CComPtr<IListViewColumns> pColumns = controls.exlvwVirtual->GetColumns();
	pColumns->Add(OLESTR("Column 1"), -1, 160, 0, alLeft, 1, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 2"), -1, 160, 0, alCenter, 2, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 3"), -1, 160, 0, alCenter, 3, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::UpdateMenu(void)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwnormalIsFocused) {
		switch(controls.exlvwNormal->GetView()) {
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
	} else {
		switch(controls.exlvwVirtual->GetView()) {
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
	}
	EnableMenuItem(hMenu, ID_VIEW_TILES, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
	EnableMenuItem(hMenu, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
	EnableMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));

	hMenu = GetSubMenu(hMainMenu, 1);
	if(exlvwnormalIsFocused) {
		switch(controls.exlvwNormal->GetItemAlignment()) {
			case iaTop:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_TOP, MF_BYCOMMAND);
				break;
			case iaLeft:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_LEFT, MF_BYCOMMAND);
				break;
		}
	} else {
		switch(controls.exlvwVirtual->GetItemAlignment()) {
			case iaTop:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_TOP, MF_BYCOMMAND);
				break;
			case iaLeft:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_LEFT, MF_BYCOMMAND);
				break;
		}
	}
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwnormal(long /*hWndHeader*/)
{
	InsertColumnsNormal();
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwvirtual(long /*hWndHeader*/)
{
	InsertColumnsVirtual();
}

void __stdcall CMainDlg::ItemGetDisplayInfoExlvwvirtual(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* /*Indent*/, long* /*groupID*/, LPSAFEARRAY* TileViewColumns, long /*maxItemTextLength*/, BSTR* itemText, long* /*OverlayIndex*/, long* /*StateImageIndex*/, long* /*itemStates*/, VARIANT_BOOL* /*dontAskAgain*/)
{
	CComQIPtr<IListViewItem> pItem = listItem;
	if(requestedInfo & riItemText) {
		CComQIPtr<IListViewSubItem> pSubItem = listSubItem;
		if(pSubItem) {
			CAtlString tmp;
			tmp.Format(TEXT("Item %i, SubItem %i"), pItem->GetIndex(), pSubItem->GetIndex());

			*itemText = _bstr_t(tmp).Detach();
		} else {
			CAtlString tmp;
			tmp.Format(TEXT("Item %i"), pItem->GetIndex());

			*itemText = _bstr_t(tmp).Detach();
		}
	}
	if(requestedInfo & riIconIndex) {
		*IconIndex = pItem->GetIndex() % 10;
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
}

void __stdcall CMainDlg::MapGroupWideToTotalItemIndexExlvwvirtual(LONG groupIndex, LONG groupWideItemIndex, LONG* totalItemIndex)
{
	*totalItemIndex = groupIndex * cItemsHalf + groupWideItemIndex;
}
