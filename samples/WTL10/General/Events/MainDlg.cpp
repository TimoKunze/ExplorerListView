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
		IDispEventImpl<IDC_EXLVWU, CMainDlg, &__uuidof(ExLVwLibU::_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>::DispEventUnadvise(controls.exlvwU);
		controls.exlvwU.Release();
	}
	if(controls.exlvwA) {
		IDispEventImpl<IDC_EXLVWA, CMainDlg, &__uuidof(ExLVwLibA::_IExplorerListViewEvents), &LIBID_ExLVwLibA, 1, 7>::DispEventUnadvise(controls.exlvwA);
		controls.exlvwA.Release();
	}

	if(!IsComctl32Version600OrNewer()) {
		DeleteObject(LongToHandle(hBMPDownArrow));
		DeleteObject(LongToHandle(hBMPUpArrow));
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
								HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
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
								HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(pIconData), SizeofResource(hMod, hResource), TRUE, 0x00030000, 0, 0, LR_DEFAULTCOLOR);
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

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	controls.aboutButton = GetDlgItem(ID_APP_ABOUT);

	exlvwUContainerWnd.SubclassWindow(GetDlgItem(IDC_EXLVWU));
	exlvwUContainerWnd.QueryControl(__uuidof(ExLVwLibU::IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwU));
	if(controls.exlvwU) {
		IDispEventImpl<IDC_EXLVWU, CMainDlg, &__uuidof(ExLVwLibU::_IExplorerListViewEvents), &LIBID_ExLVwLibU, 1, 7>::DispEventAdvise(controls.exlvwU);
		exlvwUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())));
		InsertColumnsU();

		if(!IsComctl32Version600OrNewer()) {
			CComPtr<ExLVwLibU::IListViewColumn> pColumn = controls.exlvwU->GetColumns()->GetItem(0, ExLVwLibU::citIndex);
			pColumn->PutBitmapHandle(static_cast<OLE_HANDLE>(-1));
			hBMPDownArrow = pColumn->GetBitmapHandle();
			pColumn->PutBitmapHandle(static_cast<OLE_HANDLE>(-2));
			hBMPUpArrow = pColumn->GetBitmapHandle();
			pColumn->PutBitmapHandle(0);
		}

		if(IsComctl32Version600OrNewer()) {
			InsertGroupsU();
		}
		InsertItemsU();
		InsertFooterItemsU();
	}

	exlvwAContainerWnd.SubclassWindow(GetDlgItem(IDC_EXLVWA));
	exlvwAContainerWnd.QueryControl(__uuidof(ExLVwLibA::IExplorerListView), reinterpret_cast<LPVOID*>(&controls.exlvwA));
	if(controls.exlvwA) {
		IDispEventImpl<IDC_EXLVWA, CMainDlg, &__uuidof(ExLVwLibA::_IExplorerListViewEvents), &LIBID_ExLVwLibA, 1, 7>::DispEventAdvise(controls.exlvwA);
		exlvwAWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.exlvwA->GethWnd())));
		InsertColumnsA();
		if(IsComctl32Version600OrNewer()) {
			InsertGroupsA();
		}
		InsertItemsA();
		InsertFooterItemsA();
	}

	if(CTheme::IsThemingSupported() && IsComctl32Version600OrNewer()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed && pfnIsAppThemed()) {
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), L"explorer", NULL);
				SetWindowTheme(static_cast<HWND>(LongToHandle(controls.exlvwA->GethWnd())), L"explorer", NULL);
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
	if(controls.logEdit && controls.aboutButton) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			CRect rc;
			GetClientRect(&rc);
			int cx = static_cast<int>(0.4 * static_cast<double>(rc.Width()));
			controls.logEdit.SetWindowPos(NULL, rc.Width() - cx, 0, cx, rc.Height() - 32, 0);
			controls.aboutButton.SetWindowPos(NULL, rc.Width() - cx, rc.Height() - 27, 0, 0, SWP_NOSIZE);
			exlvwUContainerWnd.SetWindowPos(NULL, 0, 0, rc.Width() - cx - 5, (rc.Height() - 5) / 2, SWP_NOMOVE);
			exlvwAContainerWnd.SetWindowPos(NULL, 0, (rc.Height() - 5) / 2 + 5, rc.Width() - cx - 5, (rc.Height() - 5) / 2, 0);
		}
	}

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		if(controls.exlvwA) {
			controls.exlvwA->About();
		}
	} else {
		if(controls.exlvwU) {
			controls.exlvwU->About();
		}
	}
	return 0;
}

LRESULT CMainDlg::OnViewIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutView(ExLVwLibA::vIcons);
	} else {
		controls.exlvwU->PutView(ExLVwLibU::vIcons);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewSmallIcons(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutView(ExLVwLibA::vSmallIcons);
	} else {
		controls.exlvwU->PutView(ExLVwLibU::vSmallIcons);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewList(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutView(ExLVwLibA::vList);
	} else {
		controls.exlvwU->PutView(ExLVwLibU::vList);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewDetails(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutView(ExLVwLibA::vDetails);
		controls.exlvwA->PutFullRowSelect(ExLVwLibA::frsExtendedMode);
	} else {
		controls.exlvwU->PutView(ExLVwLibU::vDetails);
		controls.exlvwU->PutFullRowSelect(ExLVwLibU::frsExtendedMode);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutView(ExLVwLibA::vTiles);
		controls.exlvwA->PutFullRowSelect(ExLVwLibA::frsDisabled);
	} else {
		controls.exlvwU->PutView(ExLVwLibU::vTiles);
		controls.exlvwU->PutFullRowSelect(ExLVwLibU::frsDisabled);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewExtendedTiles(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutView(ExLVwLibA::vExtendedTiles);
		controls.exlvwA->PutFullRowSelect(ExLVwLibA::frsDisabled);
	} else {
		controls.exlvwU->PutView(ExLVwLibU::vExtendedTiles);
		controls.exlvwU->PutFullRowSelect(ExLVwLibU::frsDisabled);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewAutoArrange(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutAutoArrangeItems(ExLVwLibA::aaiNone);
		} else {
			controls.exlvwA->PutAutoArrangeItems(ExLVwLibA::aaiIntelligent);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutAutoArrangeItems(ExLVwLibU::aaiNone);
		} else {
			controls.exlvwU->PutAutoArrangeItems(ExLVwLibU::aaiIntelligent);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewJustifyIconColumns(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutJustifyIconColumns(VARIANT_FALSE);
		} else {
			controls.exlvwA->PutJustifyIconColumns(VARIANT_TRUE);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutJustifyIconColumns(VARIANT_FALSE);
		} else {
			controls.exlvwU->PutJustifyIconColumns(VARIANT_TRUE);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewSnapToGrid(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutSnapToGrid(VARIANT_FALSE);
		} else {
			controls.exlvwA->PutSnapToGrid(VARIANT_TRUE);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutSnapToGrid(VARIANT_FALSE);
		} else {
			controls.exlvwU->PutSnapToGrid(VARIANT_TRUE);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewShowInGroups(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutShowGroups(VARIANT_FALSE);
		} else {
			controls.exlvwA->PutShowGroups(VARIANT_TRUE);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutShowGroups(VARIANT_FALSE);
		} else {
			controls.exlvwU->PutShowGroups(VARIANT_TRUE);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewCheckOnSelect(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutCheckItemOnSelect(VARIANT_FALSE);
		} else {
			controls.exlvwA->PutCheckItemOnSelect(VARIANT_TRUE);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutCheckItemOnSelect(VARIANT_FALSE);
		} else {
			controls.exlvwU->PutCheckItemOnSelect(VARIANT_TRUE);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewAutoSizeColumns(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutAutoSizeColumns(VARIANT_FALSE);
		} else {
			controls.exlvwA->PutAutoSizeColumns(VARIANT_TRUE);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutAutoSizeColumns(VARIANT_FALSE);
		} else {
			controls.exlvwU->PutAutoSizeColumns(VARIANT_TRUE);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewResizableColumns(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		if(GetMenuState(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwA->PutResizableColumns(VARIANT_FALSE);
		} else {
			controls.exlvwA->PutResizableColumns(VARIANT_TRUE);
		}
	} else {
		if(GetMenuState(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND) & MF_CHECKED) {
			controls.exlvwU->PutResizableColumns(VARIANT_FALSE);
		} else {
			controls.exlvwU->PutResizableColumns(VARIANT_TRUE);
		}
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnAlignTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutItemAlignment(ExLVwLibA::iaTop);
	} else {
		controls.exlvwU->PutItemAlignment(ExLVwLibU::iaTop);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnAlignLeft(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(exlvwaIsFocused) {
		controls.exlvwA->PutItemAlignment(ExLVwLibA::iaLeft);
	} else {
		controls.exlvwU->PutItemAlignment(ExLVwLibU::iaLeft);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnSetFocusU(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	exlvwaIsFocused = FALSE;
	UpdateMenu();
	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnSetFocusA(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	exlvwaIsFocused = TRUE;
	UpdateMenu();
	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::InsertColumnsA(void)
{
	CComPtr<ExLVwLibA::IListViewColumns> pColumns = controls.exlvwA->GetColumns();
	pColumns->Add(OLESTR("Column 1"), -1, 160, 0, ExLVwLibA::alLeft, 1, 1, VARIANT_TRUE, VARIANT_TRUE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 2"), -1, 160, 0, ExLVwLibA::alCenter, 2, 0, VARIANT_TRUE, VARIANT_TRUE, VARIANT_FALSE);
	pColumns->Add(OLESTR("Column 3"), -1, 160, 0, ExLVwLibA::alRight, 3, 0, VARIANT_TRUE, VARIANT_TRUE, VARIANT_FALSE);
}

void CMainDlg::InsertFooterItemsA(void)
{
	if(IsComctl32Version610OrNewer()) {
		controls.exlvwA->PuthImageList(ExLVwLibA::ilFooterItems, HandleToLong(controls.smallImageList.m_hImageList));
		CComPtr<ExLVwLibA::IListViewFooterItems> pItems = controls.exlvwA->GetFooterItems();
		pItems->Add(OLESTR("Sleep"), -1, 0, 1);
		pItems->Add(OLESTR("Rule the world"), -1, 1, 2);
		controls.exlvwA->ShowFooter();
	}
}

void CMainDlg::InsertGroupsA(void)
{
	if(IsComctl32Version610OrNewer()) {
		controls.exlvwA->PuthImageList(ExLVwLibA::ilGroups, HandleToLong(controls.smallImageList.m_hImageList));
	}
	CComPtr<ExLVwLibA::IListViewGroups> pGroups = controls.exlvwA->GetGroups();
	CComPtr<ExLVwLibA::IListViewGroup> pGroup = pGroups->Add(OLESTR("Group 1"), 1, -1, 1, ExLVwLibA::alLeft, 0, VARIANT_FALSE, VARIANT_FALSE, OLESTR("Footer 1"), ExLVwLibA::alRight, OLESTR("Subtitle"), OLESTR("Task"), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	if(IsComctl32Version610OrNewer()) {
		pGroup->PutDescriptionTextBottom(OLESTR("Bottom"));
		pGroup->PutDescriptionTextTop(OLESTR("Top"));
	}
	pGroup = pGroups->Add(OLESTR("Group 2"), 2, -1, 1, ExLVwLibA::alCenter, 1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), ExLVwLibA::alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	pGroup = pGroups->Add(OLESTR("Group 3"), 3, -1, 1, ExLVwLibA::alRight, 2, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), ExLVwLibA::alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
}

void CMainDlg::InsertItemsA(void)
{
	HBITMAP hBMP = static_cast<HBITMAP>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDB_WATERMARK2), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR));
	controls.exlvwA->PutBkImage(_variant_t(HandleToLong(hBMP)));

	controls.exlvwA->PuthImageList(ExLVwLibA::ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
	controls.exlvwA->PuthImageList(ExLVwLibA::ilLarge, HandleToLong(controls.largeImageList.m_hImageList));
	controls.exlvwA->PuthImageList(ExLVwLibA::ilExtraLarge, HandleToLong(controls.extralargeImageList.m_hImageList));

	CComPtr<ExLVwLibA::IListViewItems> pItems = controls.exlvwA->GetListItems();
	CComPtr<IRecordInfo> pRecordInfo = NULL;
	CLSID clsidTILEVIEWSUBITEM = {0};
	CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
	ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, 1, 7, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
	LPSAFEARRAY pTileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);
	ExLVwLibA::TILEVIEWSUBITEM element = {0};
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

	CComPtr<ExLVwLibA::IListViewSubItems> pSubItems = pItems->Add(OLESTR("Item 1"), -1, 0, 0, 1, 1, tileViewColumns12)->GetSubItems();
	CComPtr<ExLVwLibA::IListViewSubItem> pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 1, SubItem 1"));
	pSubItem->PutIconIndex(1);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 1, SubItem 2"));
	pSubItem->PutIconIndex(2);

	pSubItems = pItems->Add(OLESTR("Item 2"), -1, 1, 0, 2, 2, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 2, SubItem 1"));
	pSubItem->PutIconIndex(2);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 2, SubItem 2"));
	pSubItem->PutIconIndex(3);

	pSubItems = pItems->Add(OLESTR("Item 3"), -1, 2, 0, 3, 3, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 3, SubItem 1"));
	pSubItem->PutIconIndex(3);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 3, SubItem 2"));
	pSubItem->PutIconIndex(4);

	pSubItems = pItems->Add(OLESTR("Item 4"), -1, 3, 0, 4, 1, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 4, SubItem 1"));
	pSubItem->PutIconIndex(4);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 4, SubItem 2"));
	pSubItem->PutIconIndex(5);

	pSubItems = pItems->Add(OLESTR("Item 5"), -1, 4, 0, 5, 2, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 5, SubItem 1"));
	pSubItem->PutIconIndex(5);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 5, SubItem 2"));
	pSubItem->PutIconIndex(6);

	pSubItems = pItems->Add(OLESTR("Item 6"), -1, 5, 0, 6, 3, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 6, SubItem 1"));
	pSubItem->PutIconIndex(6);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 6, SubItem 2"));
	pSubItem->PutIconIndex(7);

	pSubItems = pItems->Add(OLESTR("Item 7"), -1, 6, 0, 7, 1, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 7, SubItem 1"));
	pSubItem->PutIconIndex(7);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 7, SubItem 2"));
	pSubItem->PutIconIndex(8);

	pSubItems = pItems->Add(OLESTR("Item 8"), -1, 7, 0, 8, 2, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 8, SubItem 1"));
	pSubItem->PutIconIndex(8);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 8, SubItem 2"));
	pSubItem->PutIconIndex(9);

	pSubItems = pItems->Add(OLESTR("Item 9"), -1, 8, 0, 9, 3, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 9, SubItem 1"));
	pSubItem->PutIconIndex(9);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 9, SubItem 2"));
	pSubItem->PutIconIndex(0);

	pSubItems = pItems->Add(OLESTR("Item 10"), -1, 9, 0, 10, 1, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 10, SubItem 1"));
	pSubItem->PutIconIndex(0);
	pSubItem = pSubItems->GetItem(2, ExLVwLibA::citIndex);
	pSubItem->PutText(OLESTR("Item 10, SubItem 2"));
	pSubItem->PutIconIndex(1);
}

void CMainDlg::InsertColumnsU(void)
{
	controls.exlvwU->PuthImageList(ExLVwLibU::ilHeader, HandleToLong(controls.smallImageList.m_hImageList));
	CComPtr<ExLVwLibU::IListViewColumns> pColumns = controls.exlvwU->GetColumns();
	pColumns->Add(OLESTR("Column 1"), -1, 160, 0, ExLVwLibU::alLeft, 1, 1, VARIANT_TRUE, VARIANT_TRUE, VARIANT_FALSE)->PutIconIndex(5);
	pColumns->Add(OLESTR("Column 2"), -1, 160, 0, ExLVwLibU::alCenter, 2, 0, VARIANT_TRUE, VARIANT_TRUE, VARIANT_FALSE)->PutIconIndex(6);
	pColumns->Add(OLESTR("Column 3"), -1, 160, 0, ExLVwLibU::alRight, 3, 0, VARIANT_TRUE, VARIANT_TRUE, VARIANT_FALSE)->PutIconIndex(7);
}

void CMainDlg::InsertFooterItemsU(void)
{
	if(IsComctl32Version610OrNewer()) {
		controls.exlvwU->PuthImageList(ExLVwLibU::ilFooterItems, HandleToLong(controls.smallImageList.m_hImageList));
		CComPtr<ExLVwLibU::IListViewFooterItems> pItems = controls.exlvwU->GetFooterItems();
		pItems->Add(OLESTR("Sleep"), -1, 0, 1);
		pItems->Add(OLESTR("Rule the world"), -1, 1, 2);
		controls.exlvwU->ShowFooter();
	}
}

void CMainDlg::InsertGroupsU(void)
{
	if(IsComctl32Version610OrNewer()) {
		controls.exlvwU->PuthImageList(ExLVwLibU::ilGroups, HandleToLong(controls.smallImageList.m_hImageList));
	}
	CComPtr<ExLVwLibU::IListViewGroups> pGroups = controls.exlvwU->GetGroups();
	CComPtr<ExLVwLibU::IListViewGroup> pGroup = pGroups->Add(OLESTR("Group 1"), 1, -1, 1, ExLVwLibU::alLeft, 0, VARIANT_FALSE, VARIANT_FALSE, OLESTR("Footer 1"), ExLVwLibU::alRight, OLESTR("Subtitle"), OLESTR("Task"), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	if(IsComctl32Version610OrNewer()) {
		pGroup->PutDescriptionTextBottom(OLESTR("Bottom"));
		pGroup->PutDescriptionTextTop(OLESTR("Top"));
	}
	pGroup = pGroups->Add(OLESTR("Group 2"), 2, -1, 1, ExLVwLibU::alCenter, 1, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), ExLVwLibU::alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
	pGroup = pGroups->Add(OLESTR("Group 3"), 3, -1, 1, ExLVwLibU::alRight, 2, VARIANT_FALSE, VARIANT_FALSE, OLESTR(""), ExLVwLibU::alLeft, OLESTR(""), OLESTR(""), OLESTR(""), VARIANT_FALSE, VARIANT_TRUE);
}

void CMainDlg::InsertItemsU(void)
{
	HBITMAP hBMP = static_cast<HBITMAP>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDB_WATERMARK1), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR));
	controls.exlvwU->PutBkImage(_variant_t(HandleToLong(hBMP)));

	controls.exlvwU->PuthImageList(ExLVwLibU::ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
	controls.exlvwU->PuthImageList(ExLVwLibU::ilLarge, HandleToLong(controls.largeImageList.m_hImageList));
	controls.exlvwU->PuthImageList(ExLVwLibU::ilExtraLarge, HandleToLong(controls.extralargeImageList.m_hImageList));

	CComPtr<ExLVwLibU::IListViewItems> pItems = controls.exlvwU->GetListItems();
	CComPtr<IRecordInfo> pRecordInfo = NULL;
	CLSID clsidTILEVIEWSUBITEM = {0};
	CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
	ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, 1, 7, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
	LPSAFEARRAY pTileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);
	ExLVwLibU::TILEVIEWSUBITEM element = {0};
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

	CComPtr<ExLVwLibU::IListViewSubItems> pSubItems = pItems->Add(OLESTR("Item 1"), -1, 0, 0, 1, 1, tileViewColumns12)->GetSubItems();
	CComPtr<ExLVwLibU::IListViewSubItem> pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 1, SubItem 1"));
	pSubItem->PutIconIndex(1);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 1, SubItem 2"));
	pSubItem->PutIconIndex(2);

	pSubItems = pItems->Add(OLESTR("Item 2"), -1, 1, 0, 2, 2, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 2, SubItem 1"));
	pSubItem->PutIconIndex(2);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 2, SubItem 2"));
	pSubItem->PutIconIndex(3);

	pSubItems = pItems->Add(OLESTR("Item 3"), -1, 2, 0, 3, 3, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 3, SubItem 1"));
	pSubItem->PutIconIndex(3);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 3, SubItem 2"));
	pSubItem->PutIconIndex(4);

	pSubItems = pItems->Add(OLESTR("Item 4"), -1, 3, 0, 4, 1, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 4, SubItem 1"));
	pSubItem->PutIconIndex(4);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 4, SubItem 2"));
	pSubItem->PutIconIndex(5);

	pSubItems = pItems->Add(OLESTR("Item 5"), -1, 4, 0, 5, 2, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 5, SubItem 1"));
	pSubItem->PutIconIndex(5);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 5, SubItem 2"));
	pSubItem->PutIconIndex(6);

	pSubItems = pItems->Add(OLESTR("Item 6"), -1, 5, 0, 6, 3, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 6, SubItem 1"));
	pSubItem->PutIconIndex(6);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 6, SubItem 2"));
	pSubItem->PutIconIndex(7);

	pSubItems = pItems->Add(OLESTR("Item 7"), -1, 6, 0, 7, 1, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 7, SubItem 1"));
	pSubItem->PutIconIndex(7);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 7, SubItem 2"));
	pSubItem->PutIconIndex(8);

	pSubItems = pItems->Add(OLESTR("Item 8"), -1, 7, 0, 8, 2, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 8, SubItem 1"));
	pSubItem->PutIconIndex(8);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 8, SubItem 2"));
	pSubItem->PutIconIndex(9);

	pSubItems = pItems->Add(OLESTR("Item 9"), -1, 8, 0, 9, 3, tileViewColumns12)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 9, SubItem 1"));
	pSubItem->PutIconIndex(9);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 9, SubItem 2"));
	pSubItem->PutIconIndex(0);

	pSubItems = pItems->Add(OLESTR("Item 10"), -1, 9, 0, 10, 1, tileViewColumns21)->GetSubItems();
	pSubItem = pSubItems->GetItem(1, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 10, SubItem 1"));
	pSubItem->PutIconIndex(0);
	pSubItem = pSubItems->GetItem(2, ExLVwLibU::citIndex);
	pSubItem->PutText(OLESTR("Item 10, SubItem 2"));
	pSubItem->PutIconIndex(1);
}

void CMainDlg::UpdateMenu(void)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(exlvwaIsFocused) {
		switch(controls.exlvwA->GetView()) {
			case ExLVwLibA::vDetails:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_DETAILS, MF_BYCOMMAND);
				break;
			case ExLVwLibA::vIcons:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_ICONS, MF_BYCOMMAND);
				break;
			case ExLVwLibA::vList:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_LIST, MF_BYCOMMAND);
				break;
			case ExLVwLibA::vSmallIcons:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_SMALLICONS, MF_BYCOMMAND);
				break;
			case ExLVwLibA::vTiles:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_TILES, MF_BYCOMMAND);
				break;
			case ExLVwLibA::vExtendedTiles:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND);
				break;
		}
		EnableMenuItem(hMenu, ID_VIEW_TILES, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		EnableMenuItem(hMenu, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwA->GetAutoArrangeItems() == ExLVwLibA::aaiNone) {
			CheckMenuItem(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND | MF_CHECKED);
		}
		if(controls.exlvwA->GetJustifyIconColumns() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwA->GetSnapToGrid() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwA->GetShowGroups() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwA->GetCheckItemOnSelect() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwA->GetAutoSizeColumns() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwA->GetResizableColumns() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
	} else {
		switch(controls.exlvwU->GetView()) {
			case ExLVwLibU::vDetails:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_DETAILS, MF_BYCOMMAND);
				break;
			case ExLVwLibU::vIcons:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_ICONS, MF_BYCOMMAND);
				break;
			case ExLVwLibU::vList:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_LIST, MF_BYCOMMAND);
				break;
			case ExLVwLibU::vSmallIcons:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_SMALLICONS, MF_BYCOMMAND);
				break;
			case ExLVwLibU::vTiles:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_TILES, MF_BYCOMMAND);
				break;
			case ExLVwLibU::vExtendedTiles:
				CheckMenuRadioItem(hMenu, ID_VIEW_ICONS, ID_VIEW_EXTENDEDTILES, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND);
				break;
		}
		EnableMenuItem(hMenu, ID_VIEW_TILES, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		EnableMenuItem(hMenu, ID_VIEW_EXTENDEDTILES, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwU->GetAutoArrangeItems() == ExLVwLibU::aaiNone) {
			CheckMenuItem(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND | MF_CHECKED);
		}
		if(controls.exlvwU->GetJustifyIconColumns() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_JUSTIFYICONCOLUMNS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwU->GetSnapToGrid() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwU->GetShowGroups() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_SHOWINGROUPS, MF_BYCOMMAND | (IsComctl32Version600OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwU->GetCheckItemOnSelect() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_CHECKONSELECT, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwU->GetAutoSizeColumns() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_AUTOSIZECOLUMNS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
		if(controls.exlvwU->GetResizableColumns() == VARIANT_FALSE) {
			CheckMenuItem(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND | MF_UNCHECKED);
		} else {
			CheckMenuItem(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND | MF_CHECKED);
		}
		EnableMenuItem(hMenu, ID_VIEW_RESIZABLECOLUMNS, MF_BYCOMMAND | (IsComctl32Version610OrNewer() ? MF_ENABLED : MF_DISABLED | MF_GRAYED));
	}

	hMenu = GetSubMenu(hMainMenu, 1);
	if(exlvwaIsFocused) {
		switch(controls.exlvwA->GetItemAlignment()) {
			case ExLVwLibA::iaTop:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_TOP, MF_BYCOMMAND);
				break;
			case ExLVwLibA::iaLeft:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_LEFT, MF_BYCOMMAND);
				break;
		}
	} else {
		switch(controls.exlvwU->GetItemAlignment()) {
			case ExLVwLibU::iaTop:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_TOP, MF_BYCOMMAND);
				break;
			case ExLVwLibU::iaLeft:
				CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_LEFT, MF_BYCOMMAND);
				break;
		}
	}
}

void __stdcall CMainDlg::AbortedDragExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_AbortedDrag")));

	controls.exlvwU->PutRefDropHilitedItem(NULL);
}

void __stdcall CMainDlg::AfterScrollExlvwu(long dx, long dy)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_AfterScroll: dx=%i, dy=%i"), dx, dy);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeScrollExlvwu(long dx, long dy)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_BeforeScroll: dx=%i, dy=%i"), dx, dy);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginColumnResizingExlvwu(LPDISPATCH column, VARIANT_BOOL* cancel)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_BeginColumnResizing: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_BeginColumnResizing: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancel=%i"), *cancel);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginMarqueeSelectionExlvwu(BOOL* cancel)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_BeginMarqueeSelection: cancel=%i"), *cancel);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CacheItemsHintExlvwu(LPDISPATCH firstItem, LPDISPATCH lastItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExLVwU_CacheItemsHint: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CacheItemsHint: firstItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewItem> pLastItem = lastItem;
	if(pLastItem) {
		BSTR text = pLastItem->GetText();
		str += TEXT(", lastItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", lastItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::CancelSubItemEditExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemEditModeConstants editMode)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwU_CancelSubItemEdit: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CancelSubItemEdit: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", editMode=%i"), editMode);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangedExlvwu(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExLVwU_CaretChanged: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CaretChanged: previousCaretItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangedSortOrderExlvwu(long previousSortOrder, long newSortOrder)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_ChangedSortOrder: previousSortOrder=%i, newSortOrder=%i"), previousSortOrder, newSortOrder);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangedWorkAreasExlvwu(LPDISPATCH WorkAreas)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewWorkAreas> pWorkAreas = WorkAreas;
	if(pWorkAreas) {
		str.Format(TEXT("ExLVwU_ChangedWorkAreas: WorkAreas.Count=%i"), pWorkAreas->Count());
	} else {
		str += TEXT("ExLVwU_ChangedWorkAreas: WorkAreas=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingSortOrderExlvwu(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_ChangingSortOrder: previousSortOrder=%i, newSortOrder=%i, cancelChange=%i"), previousSortOrder, newSortOrder, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingWorkAreasExlvwu(LPDISPATCH WorkAreas, VARIANT_BOOL* cancelChanges)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewWorkAreas> pWorkAreas = WorkAreas;
	if(pWorkAreas) {
		str.Format(TEXT("ExLVwU_ChangingWorkAreas: WorkAreas.Count=%i"), pWorkAreas->Count());
	} else {
		str += TEXT("ExLVwU_ChangingWorkAreas: WorkAreas=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChanges=%i"), *cancelChanges);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_Click: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_Click: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CollapsedGroupExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_CollapsedGroup: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CollapsedGroup: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnBeginDragExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDragDrop)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnBeginDrag: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnBeginDrag: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, doAutomaticDragDrop=%i"), button, shift, x, y, hitTestDetails, *doAutomaticDragDrop);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pColumn) {
		CComPtr<ExLVwLibU::IListViewColumns> pColumns = controls.exlvwU->GetColumns();

		CComQIPtr<IEnumVARIANT> pIterator = pColumns;
		_variant_t vColumn;
		while(pIterator->Next(1, &vColumn, NULL) == S_OK) {
			if(vColumn.vt == VT_DISPATCH) {
				ExLVwLibU::SortArrowConstants curArrow = ExLVwLibU::saNone;
				ExLVwLibU::SortArrowConstants newArrow = ExLVwLibU::saNone;

				CComQIPtr<ExLVwLibU::IListViewColumn> pCol = vColumn.pdispVal;
				if(pCol->GetIndex() == pColumn->GetIndex()) {
					if(IsComctl32Version600OrNewer()) {
						curArrow = pCol->GetSortArrow();
					} else {
						if(pCol->GetBitmapHandle() == 0) {
							curArrow = ExLVwLibU::saNone;
						} else if(pCol->GetBitmapHandle() == static_cast<OLE_HANDLE>(hBMPDownArrow)) {
							curArrow = ExLVwLibU::saDown;
						} else if(pCol->GetBitmapHandle() == static_cast<OLE_HANDLE>(hBMPUpArrow)) {
							curArrow = ExLVwLibU::saUp;
						}
					}

					if(curArrow == ExLVwLibU::saUp) {
						newArrow = ExLVwLibU::saDown;
					} else {
						newArrow = ExLVwLibU::saUp;
					}
				} else {
					newArrow = ExLVwLibU::saNone;
				}

				if(IsComctl32Version600OrNewer()) {
					pCol->PutSortArrow(newArrow);
					if(newArrow == ExLVwLibU::saDown) {
						controls.exlvwU->PutSortOrder(ExLVwLibU::soDescending);
						controls.exlvwU->PutRefSelectedColumn(pCol);
						controls.exlvwU->SortItems(ExLVwLibU::sobShell, ExLVwLibU::sobText, ExLVwLibU::sobNone, ExLVwLibU::sobNone, ExLVwLibU::sobNone, vColumn, VARIANT_TRUE);
					} else if(newArrow == ExLVwLibU::saUp) {
						controls.exlvwU->PutSortOrder(ExLVwLibU::soAscending);
						controls.exlvwU->PutRefSelectedColumn(pCol);
						controls.exlvwU->SortItems(ExLVwLibU::sobShell, ExLVwLibU::sobText, ExLVwLibU::sobNone, ExLVwLibU::sobNone, ExLVwLibU::sobNone, vColumn, VARIANT_TRUE);
					}
				} else {
					if(newArrow == ExLVwLibU::saNone) {
						pCol->PutBitmapHandle(0);
					} else if(newArrow == ExLVwLibU::saDown) {
						pCol->PutBitmapHandle(static_cast<OLE_HANDLE>(hBMPDownArrow));
						controls.exlvwU->PutSortOrder(ExLVwLibU::soDescending);
						controls.exlvwU->SortItems(ExLVwLibU::sobShell, ExLVwLibU::sobText, ExLVwLibU::sobNone, ExLVwLibU::sobNone, ExLVwLibU::sobNone, vColumn, VARIANT_TRUE);
					} else if(newArrow == ExLVwLibU::saUp) {
						pCol->PutBitmapHandle(static_cast<OLE_HANDLE>(hBMPUpArrow));
						controls.exlvwU->PutSortOrder(ExLVwLibU::soAscending);
						controls.exlvwU->SortItems(ExLVwLibU::sobShell, ExLVwLibU::sobText, ExLVwLibU::sobNone, ExLVwLibU::sobNone, ExLVwLibU::sobNone, vColumn, VARIANT_TRUE);
					}
				}
			}
			vColumn.Clear();
		}
	}
}

void __stdcall CMainDlg::ColumnDropDownExlvwu(LPDISPATCH column, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnDropDown: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnDropDown: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);

	if(pColumn) {
		// TODO: Display drop-down
	}
}

void __stdcall CMainDlg::ColumnEndAutoDragDropExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDrop)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnEndAutoDragDrop: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnEndAutoDragDrop: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, doAutomaticDrop=%i"), button, shift, x, y, hitTestDetails, *doAutomaticDrop);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnMouseEnterExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnMouseEnter: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnMouseEnter: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnMouseLeaveExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnMouseLeave: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnMouseLeave: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnStateImageChangedExlvwu(LPDISPATCH column, long previousStateImageIndex, long newStateImageIndex, long causedBy)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnStateImageChanged: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnStateImageChanged: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i"), previousStateImageIndex, newStateImageIndex, causedBy);
	str += tmp;

	AddLogEntry(str);

	if(pColumn) {
		if(pColumn->GetIndex() == 0) {
			CComQIPtr<ExLVwLibU::IListViewItems> pItems = controls.exlvwU->GetListItems();
			if(pItems) {
				pItems->GetItem(-1, 0, ExLVwLibU::iitIndex)->PutStateImageIndex(pColumn->GetStateImageIndex());
			}
		}
	}
}

void __stdcall CMainDlg::ColumnStateImageChangingExlvwu(LPDISPATCH column, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ColumnStateImageChanging: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ColumnStateImageChanging: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i, cancelChange=%i"), previousStateImageIndex, *newStateImageIndex, causedBy, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareGroupsExlvwu(LPDISPATCH firstGroup, LPDISPATCH secondGroup, long* result)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pFirstGroup = firstGroup;
	if(pFirstGroup) {
		BSTR text = pFirstGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_CompareGroups: firstGroup=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CompareGroups: firstGroup=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewGroup> pSecondGroup = secondGroup;
	if(pSecondGroup) {
		BSTR text = pSecondGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT(", secondGroup=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondGroup=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", result=%i"), *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsExlvwu(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExLVwU_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CompareItems: firstItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", result=%i"), *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ContextMenu: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ContextMenu: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);

	if(!pItem) {
		HMENU hMenu = GetMenu();
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), &pt);
		TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON, pt.x, pt.y, *this, NULL);
	}
}

void __stdcall CMainDlg::ConfigureSubItemControlExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemControlKindConstants controlKind, ExLVwLibU::SubItemEditModeConstants editMode, ExLVwLibU::SubItemControlConstants subItemControl, BSTR* themeAppName, BSTR* themeIDList, long* hFont, OLE_COLOR* textColor, long* pPropertyDescription, long pPropertyValue)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwU_ConfigureSubItemControl: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ConfigureSubItemControl: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	WCHAR pBuffer[1024];
	PropVariantToString(reinterpret_cast<REFPROPVARIANT>(pPropertyValue), pBuffer, 1024);
	tmp.Format(TEXT(", controlKind=%i, editMode=%i, subItemControl=%i, themeAppName=%s, themeIDList=%s, hFont=0x%X, textColor=0x%X, pPropertyDescription=0x%X, propertyValue=%s"), controlKind, editMode, subItemControl, *themeAppName, *themeIDList, *hFont, *textColor, *pPropertyDescription, pBuffer);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowExlvwu(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long hWndHeader)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_CreatedHeaderControlWindow: hWndHeader=0x%X"), hWndHeader);

	AddLogEntry(str);

	InsertColumnsU();
}

void __stdcall CMainDlg::CustomDrawExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL drawAllItems, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExLVwLibU::CustomDrawStageConstants drawStage, ExLVwLibU::CustomDrawItemStateConstants itemState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle, ExLVwLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_CustomDraw: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_CustomDraw: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", drawAllItems=%i, textColor=0x%X, textBackColor=0x%X, drawStage=0x%X, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawAllItems, *textColor, *textBackColor, drawStage, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_DblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_DblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowExlvwu(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowExlvwu(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedHeaderControlWindowExlvwu(long hWndHeader)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_DestroyedHeaderControlWindow: hWndHeader=0x%X"), hWndHeader);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveExlvwu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropExlvwu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	int dx = x - ptStartDrag.x;
	int dy = y - ptStartDrag.y;
	CComQIPtr<IEnumVARIANT> pEnum = controls.exlvwU->GetDraggedItems();
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ULONG ul = 0;
	_variant_t v;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<ExLVwLibU::IListViewItem> pItem = v.pdispVal;
		OLE_XPOS_PIXELS xPos;
		OLE_YPOS_PIXELS yPos;
		pItem->GetPosition(&xPos, &yPos);
		pItem->SetPosition(xPos + dx, yPos + dy);
	}
}

void __stdcall CMainDlg::EditChangeExlvwu()
{
	CAtlString str = TEXT("ExLVwU_EditChange - ");
	BSTR text = controls.exlvwU->GetEditText().Detach();
	str += text;
	SysFreeString(text);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditContextMenuExlvwu(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditContextMenu: button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditDblClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditGotFocusExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_EditGotFocus")));
}

void __stdcall CMainDlg::EditKeyDownExlvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyPressExlvwu(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyUpExlvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditLostFocusExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_EditLostFocus")));
}

void __stdcall CMainDlg::EditMClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMDblClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseDownExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseEnterExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseHoverExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseLeaveExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseMoveExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseUpExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseWheelExlvwu(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditMouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditRClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRDblClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditRDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditXClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXDblClickExlvwu(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EditXDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EmptyMarkupTextLinkClickExlvwu(long linkIndex, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_EmptyMarkupTextLinkClick: linkIndex=%i, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), linkIndex, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EmptyMarkupTextLinkClickExlvwa(long linkIndex, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EmptyMarkupTextLinkClick: linkIndex=%i, button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), linkIndex, button, shift, x, y, hitTestDetails);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EndColumnResizingExlvwu(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_EndColumnResizing: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_EndColumnResizing: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::EndSubItemEditExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemEditModeConstants editMode, long pPropertyKey, long pPropertyValue, VARIANT_BOOL* cancel)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwU_EndSubItemEdit: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_EndSubItemEdit: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	WCHAR pBuffer[1024];
	PSStringFromPropertyKey(reinterpret_cast<REFPROPERTYKEY>(pPropertyKey), pBuffer, 1024);
	tmp.Format(TEXT(", editMode=%i, pPropertyKey=%s"), editMode, pBuffer);
	str += tmp;
	PropVariantToString(reinterpret_cast<REFPROPVARIANT>(pPropertyValue), pBuffer, 1024);
	tmp.Format(TEXT(", pPropertyValue=%s, cancel=%i"), pBuffer, *cancel);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ExpandedGroupExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_ExpandedGroup: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ExpandedGroup: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FilterButtonClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, ExLVwLibU::RECTANGLE* filterButtonRectangle, VARIANT_BOOL* raiseFilterChanged)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_FilterButtonClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FilterButtonClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, filterButtonRectangle=(%i, %i)-(%i, %i), raiseFilterChanged=%i"), button, shift, x, y, filterButtonRectangle->Left, filterButtonRectangle->Top, filterButtonRectangle->Right, filterButtonRectangle->Bottom, *raiseFilterChanged);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FilterChangedExlvwu(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_FilterChanged: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FilterChanged: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FindVirtualItemExlvwu(LPDISPATCH itemToStartWith, long searchMode, VARIANT* searchFor, long searchDirection, VARIANT_BOOL wrapAtLastItem, LPDISPATCH* foundItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = itemToStartWith;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_FindVirtualItem: itemToStartWith=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FindVirtualItem: itemToStartWith=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", searchMode=%i, VarType(searchFor)=%i, searchDirection=%i, wrapAtLastItem=%i, foundItem="), searchMode, searchFor->vt, searchDirection, wrapAtLastItem);
	str += tmp;
	pItem = *foundItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FooterItemClickExlvwu(LPDISPATCH footerItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* removeFooterArea)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewFooterItem> pItem = footerItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_FooterItemClick: footerItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FooterItemClick: footerItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, removeFooterArea=%i"), button, shift, x, y, hitTestDetails, *removeFooterArea);
	str += tmp;

	AddLogEntry(str);

	if(pItem) {
		LONG itemID = pItem->GetItemData();
		switch(itemID) {
			case 1:
				MessageBox(TEXT("Good night!"));
				break;
			case 2:
				MessageBox(TEXT("Let's go, Pinky!"));
				break;
		}
	}
}

void __stdcall CMainDlg::FreeColumnDataExlvwu(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_FreeColumnData: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FreeColumnData: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeFooterItemDataExlvwu(LPDISPATCH footerItem, LONG itemData)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewFooterItem> pItem = footerItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_FreeFooterItemData: footerItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FreeFooterItemData: footerItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemData=%i"), itemData);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataExlvwu(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_FreeItemData: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_FreeItemData: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetSubItemControlExlvwu(LPDISPATCH listSubItem, ExLVwLibU::SubItemControlKindConstants controlKind, ExLVwLibU::SubItemEditModeConstants editMode, ExLVwLibU::SubItemControlConstants* subItemControl)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwU_GetSubItemControl: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GetSubItemControl: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", controlKind=%i, editMode=%i, subItemControl=%i"), controlKind, editMode, *subItemControl);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupAsynchronousDrawFailedExlvwu(LPDISPATCH Group, ExLVwLibU::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibU::ImageDrawingFailureReasonConstants failureReason, ExLVwLibU::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_GroupAsynchronousDrawFailed: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GroupAsynchronousDrawFailed: group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", failureReason=%i, furtherProcessing=%i, newImageToDraw=%i\n"), failureReason, *furtherProcessing, *newImageToDraw);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hImageList=0x%X\n"), imageDetails->hImageList);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hDC=0x%X\n"), imageDetails->hDC);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.IconIndex=%i\n"), imageDetails->IconIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.OverlayIndex=%i\n"), imageDetails->OverlayIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingStyle=%i\n"), imageDetails->DrawingStyle);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingEffects=%i\n"), imageDetails->DrawingEffects);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.BackColor=0x%X\n"), imageDetails->BackColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.ForeColor=0x%X\n"), imageDetails->ForeColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.EffectColor=0x%X\n"), imageDetails->EffectColor);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupCustomDrawExlvwu(LPDISPATCH Group, OLE_COLOR* textColor, ExLVwLibU::AlignmentConstants* headerAlignment, ExLVwLibU::AlignmentConstants* footerAlignment, ExLVwLibU::CustomDrawStageConstants drawStage, ExLVwLibU::CustomDrawItemStateConstants groupState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle, ExLVwLibU::RECTANGLE* textRectangle, ExLVwLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_GroupCustomDraw: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GroupCustomDraw: group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", textColor=0x%X, headerAlignment=%i, footerAlignment=%i, drawStage=0x%X, groupState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), textRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), *textColor, *headerAlignment, *footerAlignment, drawStage, groupState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, textRectangle->Left, textRectangle->Top, textRectangle->Right, textRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupGotFocusExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_GroupGotFocus: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GroupGotFocus: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupLostFocusExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_GroupLostFocus: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GroupLostFocus: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupSelectionChangedExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_GroupSelectionChanged: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GroupSelectionChanged: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupTaskLinkClickExlvwu(LPDISPATCH Group, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_GroupTaskLinkClick: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_GroupTaskLinkClick: group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderAbortedDragExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_HeaderAbortedDrag")));
}

void __stdcall CMainDlg::HeaderChevronClickExlvwu(LPDISPATCH firstOverflownColumn, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = firstOverflownColumn;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderChevronClick: firstOverflownColumn=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderChevronClick: firstOverflownColumn=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderContextMenuExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderContextMenu: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderContextMenu: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderCustomDrawExlvwu(LPDISPATCH column, ExLVwLibU::CustomDrawStageConstants drawStage, ExLVwLibU::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle, ExLVwLibU::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderCustomDraw: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderCustomDraw: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", drawStage=0x%X, columnState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, columnState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDividerDblClickExlvwu(LPDISPATCH column, VARIANT_BOOL* autoSizeColumn)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderDividerDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderDividerDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", autoSizeColumn=%i"), *autoSizeColumn);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDragMouseMoveExlvwu(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pDropTargt = *dropTarget;
	if(pDropTargt) {
		BSTR text = pDropTargt->GetCaption();
		str += TEXT("ExLVwU_HeaderDragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderDragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, xListView, yListView, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDropExlvwu(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pDropTargt = dropTarget;
	if(pDropTargt) {
		BSTR text = pDropTargt->GetCaption();
		str += TEXT("ExLVwU_HeaderDrop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderDrop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderGotFocusExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_HeaderGotFocus")));
}

void __stdcall CMainDlg::HeaderItemGetDisplayInfoExlvwu(LPDISPATCH column, long requestedInfo, long* IconIndex, long maxColumnCaptionLength, BSTR* columnCaption, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		CAtlString tmp;
		tmp.Format(TEXT("ExLVwU_HeaderItemGetDisplayInfo: column=%i"), pColumn->GetIndex());
		str += tmp;
	} else {
		str += TEXT("ExLVwU_HeaderItemGetDisplayInfo: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, IconIndex=%i, maxColumnCaptionLength=%i, columnCaption=%s, dontAskAgain=%i"), requestedInfo, *IconIndex, maxColumnCaptionLength, OLE2W(*columnCaption), *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderKeyDownExlvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_HeaderKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderKeyPressExlvwu(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_HeaderKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderKeyUpExlvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_HeaderKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderLostFocusExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_HeaderLostFocus")));
}

void __stdcall CMainDlg::HeaderMClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseDownExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseDown: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseDown: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseEnterExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseEnter: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseEnter: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseHoverExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseHover: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseHover: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseLeaveExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseLeave: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseLeave: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseMoveExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseMove: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseMove: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseUpExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseUp: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseUp: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseWheelExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderMouseWheel: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderMouseWheel: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLECompleteDragExlvwu(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibU::IListViewColumn> pDropTarget = dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i"), button, shift, x, y, xListView, yListView, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
		pDraggedPicture = pData->GetData(CF_DIB, -1, 1);
		_variant_t bk;
		bk.vt = VT_DISPATCH;
		if(pDraggedPicture) {
			pDraggedPicture->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&bk.pdispVal));
		} else {
			bk.pdispVal = NULL;
		}
		controls.exlvwU->PutBkImage(bk);
	}
	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.exlvwU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::HeaderOLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibU::IListViewColumn> pDropTarget = *dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, xListView, yListView, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragEnterPotentialTargetExlvwu(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_HeaderOLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragLeaveExlvwu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<ExLVwLibU::IListViewColumn> pDropTarget = dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i"), button, shift, x, y, xListView, yListView, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragLeavePotentialTargetExlvwu(void)
{
	AddLogEntry(TEXT("ExLVwU_HeaderOLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::HeaderOLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibU::IListViewColumn> pDropTarget = *dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, xListView, yListView, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEGiveFeedbackExlvwu(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_HeaderOLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEQueryContinueDragExlvwu(BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_HeaderOLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEReceivedNewDataExlvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLESetDataExlvwu(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, Index=%i, dataOrViewAspect=%i"), formatID, Index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEStartDragExlvwu(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_HeaderOLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_HeaderOLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOwnerDrawItemExlvwu(LPDISPATCH column, ExLVwLibU::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderOwnerDrawItem: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderOwnerDrawItem: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", columnState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), columnState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderRClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderRClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderRClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderRDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderRDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderRDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderXClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderXClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderXClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderXDblClickExlvwu(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_HeaderXDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HeaderXDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HotItemChangedExlvwu(LPDISPATCH previousHotItem, LPDISPATCH newHotItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pPrevItem = previousHotItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExLVwU_HotItemChanged: previousHotItem=");
		str += text;
			SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HotItemChanged: previousHotItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewItem> pNewItem = newHotItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newHotItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::HotItemChangingExlvwu(LPDISPATCH previousHotItem, LPDISPATCH newHotItem, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pPrevItem = previousHotItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExLVwU_HotItemChanging: previousHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_HotItemChanging: previousHotItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewItem> pNewItem = newHotItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newHotItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChange=%i"), *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::IncrementalSearchingExlvwu(BSTR currentSearchString, LONG* itemToSelect)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_IncrementalSearching: currentSearchString=%s, itemToSelect=%i"), OLE2W(currentSearchString), *itemToSelect);

	AddLogEntry(str);
}

void __stdcall CMainDlg::IncrementalSearchStringChangingExlvwu(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_IncrementalSearchStringChanging: currentSearchString=%s, keyCodeOfCharToBeAdded=%i, cancelChange=%i"), OLE2W(currentSearchString), keyCodeOfCharToBeAdded, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedColumnExlvwu(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_InsertedColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InsertedColumn: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedGroupExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_InsertedGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InsertedGroup: Group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemExlvwu(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_InsertedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InsertedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingColumnExlvwu(LPDISPATCH column, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IVirtualListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_InsertingColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InsertingColumn: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingGroupExlvwu(LPDISPATCH Group, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IVirtualListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_InsertingGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InsertingGroup: Group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemExlvwu(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IVirtualListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_InsertingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InsertingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::InvokeVerbFromSubItemControlExlvwu(LPDISPATCH listItem, BSTR verb)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_InvokeVerbFromSubItemControl: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_InvokeVerbFromSubItemControl: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", verb=%s"), verb);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemActivateExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemActivate: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemActivate: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", shift=%i, x=%i, y=%i"), shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemAsynchronousDrawFailedExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, ExLVwLibU::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibU::ImageDrawingFailureReasonConstants failureReason, ExLVwLibU::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemAsynchronousDrawFailed: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemAsynchronousDrawFailed: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", failureReason=%i, furtherProcessing=%i, newImageToDraw=%i\n"), failureReason, *furtherProcessing, *newImageToDraw);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hImageList=0x%X\n"), imageDetails->hImageList);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hDC=0x%X\n"), imageDetails->hDC);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.IconIndex=%i\n"), imageDetails->IconIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.OverlayIndex=%i\n"), imageDetails->OverlayIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingStyle=%i\n"), imageDetails->DrawingStyle);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingEffects=%i\n"), imageDetails->DrawingEffects);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.BackColor=0x%X\n"), imageDetails->BackColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.ForeColor=0x%X\n"), imageDetails->ForeColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.EffectColor=0x%X\n"), imageDetails->EffectColor);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemBeginDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemBeginDrag: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	CComPtr<ExLVwLibU::IListViewItems> pItems = controls.exlvwU->GetListItems();
	pItems->PutFilterType(ExLVwLibU::fpSelected, ExLVwLibU::ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(true));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
	pItems->PutFilter(ExLVwLibU::fpSelected, filter);

	ptStartDrag.x = x;
	ptStartDrag.y = y;
	controls.exlvwU->BeginDrag(controls.exlvwU->CreateItemContainer(_variant_t(pItems)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
}

void __stdcall CMainDlg::ItemBeginRDragExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemBeginRDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemBeginRDrag: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* Indent, long* groupID, LPSAFEARRAY* TileViewColumns, long maxItemTextLength, BSTR* itemText, long* OverlayIndex, long* StateImageIndex, long* itemStates, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		CAtlString tmp;
		tmp.Format(TEXT("ExLVwU_ItemGetDisplayInfo: listItem=%i"), pItem->GetIndex());
		str += tmp;
	} else {
		str += TEXT("ExLVwU_ItemGetDisplayInfo: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		CAtlString tmp;
		tmp.Format(TEXT(", listItem=%i"), pSubItem->GetIndex());
		str += tmp;
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, IconIndex=%i, Indent=%i, groupID=%i, maxItemTextLength=%i, itemText=%s, OverlayIndex=%i, StateImageIndex=%i, itemStates=0x%X, dontAskAgain=%i"), requestedInfo, *IconIndex, *Indent, *groupID, maxItemTextLength, OLE2W(*itemText), *OverlayIndex, *StateImageIndex, *itemStates, *dontAskAgain);
	str += tmp;

	AddLogEntry(str);

	if(requestedInfo & ExLVwLibU::riItemText) {
		if(pSubItem && pSubItem->GetIndex() != 0) {
			CAtlString str;
			str.Format(TEXT("Item %i, SubItem %i"), pItem->GetIndex(), pSubItem->GetIndex());
			*itemText = _bstr_t(str).Detach();
		} else {
			CAtlString str;
			str.Format(TEXT("Item %i"), pItem->GetIndex());
			*itemText = _bstr_t(str).Detach();
		}
	}
	if(requestedInfo & ExLVwLibU::riIconIndex) {
		*IconIndex = 0;
	}
	if(requestedInfo & ExLVwLibU::riTileViewColumns) {
		CComPtr<IRecordInfo> pRecordInfo = NULL;
		CLSID clsidTILEVIEWSUBITEM = {0};
		CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
		ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, 1, 7, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
		*TileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);

		ExLVwLibU::TILEVIEWSUBITEM element = {0};
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

void __stdcall CMainDlg::ItemGetGroupExlvwu(LONG itemIndex, LONG occurrenceIndex, LONG* groupIndex)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_ItemGetGroup: itemIndex=%i, occurrenceIndex=%i, groupIndex=%i"), itemIndex, occurrenceIndex, *groupIndex);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetInfoTipTextExlvwu(LPDISPATCH listItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemGetInfoTipText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemGetInfoTipText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, OLE2W(*infoTipText), *abortToolTip);
	str += tmp;

	AddLogEntry(str);

	if(controls.exlvwU->GetVirtualMode() == VARIANT_FALSE) {
		if(wcslen(OLE2W(*infoTipText)) == 0) {
			CAtlString tmp;
			tmp.Format(TEXT("ID: %i\r\nItemData: 0x%X"), pItem->GetID(), pItem->GetItemData());

			*infoTipText = _bstr_t(tmp).Detach();
		} else {
			CAtlString tmp;
			tmp.Format(TEXT("%s\r\nID: %i\r\nItemData: 0x%X"), OLE2W(*infoTipText), pItem->GetID(), pItem->GetItemData());

			*infoTipText = _bstr_t(tmp).Detach();
		}
	} else {
		*infoTipText = SysAllocString(OLESTR("Hello world!"));
	}
}

void __stdcall CMainDlg::ItemGetOccurrencesCountExlvwu(LONG itemIndex, LONG* occurrencesCount)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_ItemGetOccurrencesCount: itemIndex=%i, occurrenceIndex=%i"), itemIndex, *occurrencesCount);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterExlvwu(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemMouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemMouseEnter: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveExlvwu(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemMouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemMouseLeave: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSelectionChangedExlvwu(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemSelectionChanged: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemSelectionChanged: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSetTextExlvwu(LPDISPATCH listItem, BSTR itemText)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemSetText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemSetText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemText=%s"), OLE2W(itemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangedExlvwu(LPDISPATCH listItem, long previousStateImageIndex, long newStateImageIndex, long causedBy)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemStateImageChanged: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemStateImageChanged: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i"), previousStateImageIndex, newStateImageIndex, causedBy);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangingExlvwu(LPDISPATCH listItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_ItemStateImageChanging: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ItemStateImageChanging: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i, cancelChange=%i"), previousStateImageIndex, *newStateImageIndex, causedBy, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownExlvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);

	if(*keyCode == VK_F2) {
		controls.exlvwU->GetCaretItem(VARIANT_FALSE)->StartLabelEditing();
	} else if(*keyCode == VK_ESCAPE) {
		if(controls.exlvwU->GetDraggedItems()) {
			controls.exlvwU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::KeyPressExlvwu(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpExlvwu(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MapGroupWideToTotalItemIndexExlvwu(LONG groupIndex, LONG groupWideItemIndex, LONG* totalItemIndex)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_MapGroupWideToTotalItemIndex: groupIndex=%i, groupWideItemIndex=%i, totalItemIndex=%i"), groupIndex, groupWideItemIndex, *totalItemIndex);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MDblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseDown: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseDown: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseEnter: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseHover: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseHover: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseLeave: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseMove: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseMove: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseUp: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseUp: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(button == 1/*vbLeftButton*/) {
		if(controls.exlvwU->GetDraggedItems()) {
			// Are we within the client area?
			if(((ExLVwLibU::htAbove | ExLVwLibU::htBelow | ExLVwLibU::htToLeft | ExLVwLibU::htToRight) & hitTestDetails) == 0) {
				// yes
				controls.exlvwU->EndDrag(VARIANT_FALSE);
			} else {
				// no
				controls.exlvwU->EndDrag(VARIANT_TRUE);
			}
		}
	}
}

void __stdcall CMainDlg::MouseWheelExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_MouseWheel: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_MouseWheel: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragExlvwu(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			MessageBox(str, TEXT("Dragged files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibU::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
		pDraggedPicture = pData->GetData(CF_DIB, -1, 1);
		_variant_t bk;
		bk.vt = VT_DISPATCH;
		if(pDraggedPicture) {
			pDraggedPicture->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&bk.pdispVal));
		} else {
			bk.pdispVal = NULL;
		}
		controls.exlvwU->PutBkImage(bk);
	}
	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.exlvwU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibU::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetExlvwu(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveExlvwu(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetExlvwu(void)
{
	AddLogEntry(TEXT("ExLVwU_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibU::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackExlvwu(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragExlvwu(BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataExlvwu(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataExlvwu(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, Index=%i, dataOrViewAspect=%i"), formatID, Index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragExlvwu(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwU_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwU_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OwnerDrawItemExlvwu(LPDISPATCH listItem, ExLVwLibU::OwnerDrawItemStateConstants itemState, long hDC, ExLVwLibU::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_OwnerDrawItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_OwnerDrawItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_RClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_RDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RDblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExLVwU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	if(IsComctl32Version600OrNewer()) {
		InsertGroupsU();
	}
	InsertItemsU();
	InsertFooterItemsU();
}

void __stdcall CMainDlg::RemovedColumnExlvwu(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IVirtualListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_RemovedColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RemovedColumn: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovedGroupExlvwu(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IVirtualListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_RemovedGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RemovedGroup: Group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovedItemExlvwu(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IVirtualListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_RemovedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RemovedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingColumnExlvwu(LPDISPATCH column, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_RemovingColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RemovingColumn: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingGroupExlvwu(LPDISPATCH Group, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibU::gcHeader).Detach();
		str += TEXT("ExLVwU_RemovingGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RemovingGroup: Group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemExlvwu(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_RemovingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RemovingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamedItemExlvwu(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_RenamedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RenamedItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s"), OLE2W(previousItemText), OLE2W(newItemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamingItemExlvwu(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_RenamingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_RenamingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s, cancelRenaming=%i"), OLE2W(previousItemText), OLE2W(newItemText), *cancelRenaming);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowExlvwu()
{
	AddLogEntry(CAtlString(TEXT("ExLVwU_ResizedControlWindow")));
}

void __stdcall CMainDlg::ResizingColumnExlvwu(LPDISPATCH column, long* newColumnWidth, VARIANT_BOOL* abortResizing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwU_ResizingColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_ResizingColumn: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", newColumnWidth=%i, abortResizing=%i"), *newColumnWidth, *abortResizing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectedItemRangeExlvwu(LPDISPATCH firstItem, LPDISPATCH lastItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExLVwU_SelectedItemRange: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_SelectedItemRange: firstItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewItem> pLastItem = lastItem;
	if(pLastItem) {
		BSTR text = pLastItem->GetText();
		str += TEXT(", lastItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", lastItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::SettingItemInfoTipTextExlvwu(LPDISPATCH listItem, BSTR* infoTipText, VARIANT_BOOL* abortInfoTip)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_SettingItemInfoTipText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_SettingItemInfoTipText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", infoTipText=%s, abortInfoTip=%i"), *infoTipText, *abortInfoTip);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::StartingLabelEditingExlvwu(LPDISPATCH listItem, VARIANT_BOOL* cancelEditing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_StartingLabelEditing: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_StartingLabelEditing: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelEditing=%i"), *cancelEditing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SubItemMouseEnterExlvwu(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT("ExLVwU_SubItemMouseEnter: listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_SubItemMouseEnter: listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SubItemMouseLeaveExlvwu(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT("ExLVwU_SubItemMouseLeave: listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_SubItemMouseLeave: listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_XClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_XClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibU::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwU_XDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwU_XDblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibU::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}


void __stdcall CMainDlg::AbortedDragExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_AbortedDrag")));

	controls.exlvwA->PutRefDropHilitedItem(NULL);
}

void __stdcall CMainDlg::AfterScrollExlvwa(long dx, long dy)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_AfterScroll: dx=%i, dy=%i"), dx, dy);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeforeScrollExlvwa(long dx, long dy)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_BeforeScroll: dx=%i, dy=%i"), dx, dy);

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginColumnResizingExlvwa(LPDISPATCH column, VARIANT_BOOL* cancel)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_BeginColumnResizing: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_BeginColumnResizing: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancel=%i"), *cancel);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::BeginMarqueeSelectionExlvwa(BOOL* cancel)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_BeginMarqueeSelection: cancel=%i"), *cancel);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CacheItemsHintExlvwa(LPDISPATCH firstItem, LPDISPATCH lastItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExLVwA_CacheItemsHint: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CacheItemsHint: firstItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewItem> pLastItem = lastItem;
	if(pLastItem) {
		BSTR text = pLastItem->GetText();
		str += TEXT(", lastItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", lastItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::CancelSubItemEditExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemEditModeConstants editMode)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwA_CancelSubItemEdit: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CancelSubItemEdit: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", editMode=%i"), editMode);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CaretChangedExlvwa(LPDISPATCH previousCaretItem, LPDISPATCH newCaretItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pPrevItem = previousCaretItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExLVwA_CaretChanged: previousCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CaretChanged: previousCaretItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewItem> pNewItem = newCaretItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newCaretItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newCaretItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangedSortOrderExlvwa(long previousSortOrder, long newSortOrder)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_ChangedSortOrder: previousSortOrder=%i, newSortOrder=%i"), previousSortOrder, newSortOrder);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangedWorkAreasExlvwa(LPDISPATCH WorkAreas)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewWorkAreas> pWorkAreas = WorkAreas;
	if(pWorkAreas) {
		str.Format(TEXT("ExLVwA_ChangedWorkAreas: WorkAreas.Count=%i"), pWorkAreas->Count());
	} else {
		str += TEXT("ExLVwA_ChangedWorkAreas: WorkAreas=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingSortOrderExlvwa(long previousSortOrder, long newSortOrder, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_ChangingSortOrder: previousSortOrder=%i, newSortOrder=%i, cancelChange=%i"), previousSortOrder, newSortOrder, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ChangingWorkAreasExlvwa(LPDISPATCH WorkAreas, VARIANT_BOOL* cancelChanges)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewWorkAreas> pWorkAreas = WorkAreas;
	if(pWorkAreas) {
		str.Format(TEXT("ExLVwA_ChangingWorkAreas: WorkAreas.Count=%i"), pWorkAreas->Count());
	} else {
		str += TEXT("ExLVwA_ChangingWorkAreas: WorkAreas=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChanges=%i"), *cancelChanges);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_Click: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_Click: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CollapsedGroupExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_CollapsedGroup: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CollapsedGroup: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnBeginDragExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDragDrop)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnBeginDrag: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnBeginDrag: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, doAutomaticDragDrop=%i"), button, shift, x, y, hitTestDetails, *doAutomaticDragDrop);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pColumn) {
		CComPtr<ExLVwLibA::IListViewColumns> pColumns = controls.exlvwA->GetColumns();

		CComQIPtr<IEnumVARIANT> pIterator = pColumns;
		_variant_t vColumn;
		while(pIterator->Next(1, &vColumn, NULL) == S_OK) {
			if(vColumn.vt == VT_DISPATCH) {
				ExLVwLibA::SortArrowConstants curArrow = ExLVwLibA::saNone;
				ExLVwLibA::SortArrowConstants newArrow = ExLVwLibA::saNone;

				CComQIPtr<ExLVwLibA::IListViewColumn> pCol = vColumn.pdispVal;
				if(pCol->GetIndex() == pColumn->GetIndex()) {
					if(IsComctl32Version600OrNewer()) {
						curArrow = pCol->GetSortArrow();
					} else {
						if(pCol->GetBitmapHandle() == 0) {
							curArrow = ExLVwLibA::saNone;
						} else if(pCol->GetBitmapHandle() == static_cast<OLE_HANDLE>(hBMPDownArrow)) {
							curArrow = ExLVwLibA::saDown;
						} else if(pCol->GetBitmapHandle() == static_cast<OLE_HANDLE>(hBMPUpArrow)) {
							curArrow = ExLVwLibA::saUp;
						}
					}

					if(curArrow == ExLVwLibA::saUp) {
						newArrow = ExLVwLibA::saDown;
					} else {
						newArrow = ExLVwLibA::saUp;
					}
				} else {
					newArrow = ExLVwLibA::saNone;
				}

				if(IsComctl32Version600OrNewer()) {
					pCol->PutSortArrow(newArrow);
					if(newArrow == ExLVwLibA::saDown) {
						controls.exlvwA->PutSortOrder(ExLVwLibA::soDescending);
						controls.exlvwA->PutRefSelectedColumn(pCol);
						controls.exlvwA->SortItems(ExLVwLibA::sobShell, ExLVwLibA::sobText, ExLVwLibA::sobNone, ExLVwLibA::sobNone, ExLVwLibA::sobNone, vColumn, VARIANT_TRUE);
					} else if(newArrow == ExLVwLibA::saUp) {
						controls.exlvwA->PutSortOrder(ExLVwLibA::soAscending);
						controls.exlvwA->PutRefSelectedColumn(pCol);
						controls.exlvwA->SortItems(ExLVwLibA::sobShell, ExLVwLibA::sobText, ExLVwLibA::sobNone, ExLVwLibA::sobNone, ExLVwLibA::sobNone, vColumn, VARIANT_TRUE);
					}
				} else {
					if(newArrow == ExLVwLibA::saNone) {
						pCol->PutBitmapHandle(0);
					} else if(newArrow == ExLVwLibA::saDown) {
						pCol->PutBitmapHandle(static_cast<OLE_HANDLE>(hBMPDownArrow));
						controls.exlvwA->PutSortOrder(ExLVwLibA::soDescending);
						controls.exlvwA->SortItems(ExLVwLibA::sobShell, ExLVwLibA::sobText, ExLVwLibA::sobNone, ExLVwLibA::sobNone, ExLVwLibA::sobNone, vColumn, VARIANT_TRUE);
					} else if(newArrow == ExLVwLibA::saUp) {
						pCol->PutBitmapHandle(static_cast<OLE_HANDLE>(hBMPUpArrow));
						controls.exlvwA->PutSortOrder(ExLVwLibA::soAscending);
						controls.exlvwA->SortItems(ExLVwLibA::sobShell, ExLVwLibA::sobText, ExLVwLibA::sobNone, ExLVwLibA::sobNone, ExLVwLibA::sobNone, vColumn, VARIANT_TRUE);
					}
				}
			}
			vColumn.Clear();
		}
	}
}

void __stdcall CMainDlg::ColumnDropDownExlvwa(LPDISPATCH column, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnDropDown: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnDropDown: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);

	if(pColumn) {
		// TODO: Display drop-down
	}
}

void __stdcall CMainDlg::ColumnEndAutoDragDropExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* doAutomaticDrop)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnEndAutoDragDrop: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnEndAutoDragDrop: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, doAutomaticDrop=%i"), button, shift, x, y, hitTestDetails, *doAutomaticDrop);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnMouseEnterExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnMouseEnter: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnMouseEnter: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnMouseLeaveExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnMouseLeave: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnMouseLeave: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ColumnStateImageChangedExlvwa(LPDISPATCH column, long previousStateImageIndex, long newStateImageIndex, long causedBy)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnStateImageChanged: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnStateImageChanged: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i"), previousStateImageIndex, newStateImageIndex, causedBy);
	str += tmp;

	AddLogEntry(str);

	if(pColumn) {
		if(pColumn->GetIndex() == 0) {
			CComQIPtr<ExLVwLibA::IListViewItems> pItems = controls.exlvwA->GetListItems();
			if(pItems) {
				pItems->GetItem(-1, 0, ExLVwLibA::iitIndex)->PutStateImageIndex(pColumn->GetStateImageIndex());
			}
		}
	}
}

void __stdcall CMainDlg::ColumnStateImageChangingExlvwa(LPDISPATCH column, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ColumnStateImageChanging: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ColumnStateImageChanging: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i, cancelChange=%i"), previousStateImageIndex, *newStateImageIndex, causedBy, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareGroupsExlvwa(LPDISPATCH firstGroup, LPDISPATCH secondGroup, long* result)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pFirstGroup = firstGroup;
	if(pFirstGroup) {
		BSTR text = pFirstGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_CompareGroups: firstGroup=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CompareGroups: firstGroup=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewGroup> pSecondGroup = secondGroup;
	if(pSecondGroup) {
		BSTR text = pSecondGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT(", secondGroup=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondGroup=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", result=%i"), *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CompareItemsExlvwa(LPDISPATCH firstItem, LPDISPATCH secondItem, long* result)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExLVwA_CompareItems: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CompareItems: firstItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewItem> pSecondItem = secondItem;
	if(pSecondItem) {
		BSTR text = pSecondItem->GetText();
		str += TEXT(", secondItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", secondItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", result=%i"), *result);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ContextMenu: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ContextMenu: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);

	if(!pItem) {
		HMENU hMenu = GetMenu();
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.exlvwA->GethWnd())), &pt);
		TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON, pt.x, pt.y, *this, NULL);
	}
}

void __stdcall CMainDlg::ConfigureSubItemControlExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemControlKindConstants controlKind, ExLVwLibA::SubItemEditModeConstants editMode, ExLVwLibA::SubItemControlConstants subItemControl, BSTR* themeAppName, BSTR* themeIDList, long* hFont, OLE_COLOR* textColor, long* pPropertyDescription, long pPropertyValue)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwA_ConfigureSubItemControl: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ConfigureSubItemControl: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	WCHAR pBuffer[1024];
	PropVariantToString(reinterpret_cast<REFPROPVARIANT>(pPropertyValue), pBuffer, 1024);
	tmp.Format(TEXT(", controlKind=%i, editMode=%i, subItemControl=%i, themeAppName=%s, themeIDList=%s, hFont=0x%X, textColor=0x%X, pPropertyDescription=0x%X, propertyValue=%s"), controlKind, editMode, subItemControl, *themeAppName, *themeIDList, *hFont, *textColor, *pPropertyDescription, pBuffer);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedEditControlWindowExlvwa(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_CreatedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwa(long hWndHeader)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_CreatedHeaderControlWindow: hWndHeader=0x%X"), hWndHeader);

	AddLogEntry(str);

	InsertColumnsA();
}

void __stdcall CMainDlg::CustomDrawExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, VARIANT_BOOL drawAllItems, OLE_COLOR* textColor, OLE_COLOR* textBackColor, ExLVwLibA::CustomDrawStageConstants drawStage, ExLVwLibA::CustomDrawItemStateConstants itemState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle, ExLVwLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_CustomDraw: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_CustomDraw: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", drawAllItems=%i, textColor=0x%X, textBackColor=0x%X, drawStage=0x%X, itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawAllItems, *textColor, *textBackColor, drawStage, itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_DblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_DblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowExlvwa(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedEditControlWindowExlvwa(long hWndEdit)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_DestroyedEditControlWindow: hWndEdit=0x%X"), hWndEdit);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedHeaderControlWindowExlvwa(long hWndHeader)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_DestroyedHeaderControlWindow: hWndHeader=0x%X"), hWndHeader);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DragMouseMoveExlvwa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_DragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_DragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DropExlvwa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_Drop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_Drop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	int dx = x - ptStartDrag.x;
	int dy = y - ptStartDrag.y;
	CComQIPtr<IEnumVARIANT> pEnum = controls.exlvwA->GetDraggedItems();
	pEnum->Reset();     // NOTE: It's important to call Reset()!
	ULONG ul = 0;
	_variant_t v;
	v.Clear();
	while(pEnum->Next(1, &v, &ul) == S_OK) {
		CComQIPtr<ExLVwLibA::IListViewItem> pItem = v.pdispVal;
		OLE_XPOS_PIXELS xPos;
		OLE_YPOS_PIXELS yPos;
		pItem->GetPosition(&xPos, &yPos);
		pItem->SetPosition(xPos + dx, yPos + dy);
	}
}

void __stdcall CMainDlg::EditChangeExlvwa()
{
	CAtlString str = TEXT("ExLVwA_EditChange - ");
	BSTR text = controls.exlvwA->GetEditText().Detach();
	str += text;
	SysFreeString(text);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditContextMenuExlvwa(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditContextMenu: button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditDblClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditGotFocusExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_EditGotFocus")));
}

void __stdcall CMainDlg::EditKeyDownExlvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyPressExlvwa(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditKeyUpExlvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditLostFocusExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_EditLostFocus")));
}

void __stdcall CMainDlg::EditMClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMDblClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseDownExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseEnterExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseHoverExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseLeaveExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseMoveExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseUpExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditMouseWheelExlvwa(short button, short shift, long x, long y, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditMouseWheel: button=%i, shift=%i, x=%i, y=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, scrollAxis, wheelDelta);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditRClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditRDblClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditRDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditXClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EditXDblClickExlvwa(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_EditXDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::EndColumnResizingExlvwa(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_EndColumnResizing: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_EndColumnResizing: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::EndSubItemEditExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemEditModeConstants editMode, long pPropertyKey, long pPropertyValue, VARIANT_BOOL* cancel)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwA_EndSubItemEdit: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_EndSubItemEdit: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	WCHAR pBuffer[1024];
	PSStringFromPropertyKey(reinterpret_cast<REFPROPERTYKEY>(pPropertyKey), pBuffer, 1024);
	tmp.Format(TEXT(", editMode=%i, pPropertyKey=%s"), editMode, pBuffer);
	str += tmp;
	PropVariantToString(reinterpret_cast<REFPROPVARIANT>(pPropertyValue), pBuffer, 1024);
	tmp.Format(TEXT(", pPropertyValue=%s, cancel=%i"), pBuffer, *cancel);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ExpandedGroupExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_ExpandedGroup: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ExpandedGroup: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FilterButtonClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, ExLVwLibA::RECTANGLE* filterButtonRectangle, VARIANT_BOOL* raiseFilterChanged)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_FilterButtonClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FilterButtonClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, filterButtonRectangle=(%i, %i)-(%i, %i), raiseFilterChanged=%i"), button, shift, x, y, filterButtonRectangle->Left, filterButtonRectangle->Top, filterButtonRectangle->Right, filterButtonRectangle->Bottom, *raiseFilterChanged);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FilterChangedExlvwa(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_FilterChanged: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FilterChanged: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FindVirtualItemExlvwa(LPDISPATCH itemToStartWith, long searchMode, VARIANT* searchFor, long searchDirection, VARIANT_BOOL wrapAtLastItem, LPDISPATCH* foundItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = itemToStartWith;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_FindVirtualItem: itemToStartWith=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FindVirtualItem: itemToStartWith=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", searchMode=%i, VarType(searchFor)=%i, searchDirection=%i, wrapAtLastItem=%i, foundItem="), searchMode, searchFor->vt, searchDirection, wrapAtLastItem);
	str += tmp;
	pItem = *foundItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FooterItemClickExlvwa(LPDISPATCH footerItem, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* removeFooterArea)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewFooterItem> pItem = footerItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_FooterItemClick: footerItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FooterItemClick: footerItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, removeFooterArea=%i"), button, shift, x, y, hitTestDetails, *removeFooterArea);
	str += tmp;

	AddLogEntry(str);

	if(pItem) {
		LONG itemID = pItem->GetItemData();
		switch(itemID) {
			case 1:
				MessageBox(TEXT("Good night!"));
				break;
			case 2:
				MessageBox(TEXT("Let's go, Pinky!"));
				break;
		}
	}
}

void __stdcall CMainDlg::FreeColumnDataExlvwa(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_FreeColumnData: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FreeColumnData: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeFooterItemDataExlvwa(LPDISPATCH footerItem, LONG itemData)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewFooterItem> pItem = footerItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_FreeFooterItemData: footerItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FreeFooterItemData: footerItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemData=%i"), itemData);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::FreeItemDataExlvwa(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_FreeItemData: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_FreeItemData: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GetSubItemControlExlvwa(LPDISPATCH listSubItem, ExLVwLibA::SubItemControlKindConstants controlKind, ExLVwLibA::SubItemEditModeConstants editMode, ExLVwLibA::SubItemControlConstants* subItemControl)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetParentItem()->GetText();
		str += TEXT("ExLVwA_GetSubItemControl: listItem=");
		str += text;
		SysFreeString(text);
		text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GetSubItemControl: listItem=NULL");
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", controlKind=%i, editMode=%i, subItemControl=%i"), controlKind, editMode, *subItemControl);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupAsynchronousDrawFailedExlvwa(LPDISPATCH Group, ExLVwLibA::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibA::ImageDrawingFailureReasonConstants failureReason, ExLVwLibA::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_GroupAsynchronousDrawFailed: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GroupAsynchronousDrawFailed: group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", failureReason=%i, furtherProcessing=%i, newImageToDraw=%i\n"), failureReason, *furtherProcessing, *newImageToDraw);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hImageList=0x%X\n"), imageDetails->hImageList);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hDC=0x%X\n"), imageDetails->hDC);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.IconIndex=%i\n"), imageDetails->IconIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.OverlayIndex=%i\n"), imageDetails->OverlayIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingStyle=%i\n"), imageDetails->DrawingStyle);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingEffects=%i\n"), imageDetails->DrawingEffects);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.BackColor=0x%X\n"), imageDetails->BackColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.ForeColor=0x%X\n"), imageDetails->ForeColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.EffectColor=0x%X\n"), imageDetails->EffectColor);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupCustomDrawExlvwa(LPDISPATCH Group, OLE_COLOR* textColor, ExLVwLibA::AlignmentConstants* headerAlignment, ExLVwLibA::AlignmentConstants* footerAlignment, ExLVwLibA::CustomDrawStageConstants drawStage, ExLVwLibA::CustomDrawItemStateConstants groupState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle, ExLVwLibA::RECTANGLE* textRectangle, ExLVwLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_GroupCustomDraw: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GroupCustomDraw: group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", textColor=0x%X, headerAlignment=%i, footerAlignment=%i, drawStage=0x%X, groupState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), textRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), *textColor, *headerAlignment, *footerAlignment, drawStage, groupState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, textRectangle->Left, textRectangle->Top, textRectangle->Right, textRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupGotFocusExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_GroupGotFocus: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GroupGotFocus: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupLostFocusExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_GroupLostFocus: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GroupLostFocus: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupSelectionChangedExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_GroupSelectionChanged: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GroupSelectionChanged: group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::GroupTaskLinkClickExlvwa(LPDISPATCH Group, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_GroupTaskLinkClick: group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_GroupTaskLinkClick: group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderAbortedDragExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_HeaderAbortedDrag")));
}

void __stdcall CMainDlg::HeaderChevronClickExlvwa(LPDISPATCH firstOverflownColumn, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = firstOverflownColumn;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderChevronClick: firstOverflownColumn=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderChevronClick: firstOverflownColumn=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, showDefaultMenu=%i"), button, shift, x, y, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderContextMenuExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, VARIANT_BOOL* showDefaultMenu)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderContextMenu: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderContextMenu: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, showDefaultMenu=%i"), button, shift, x, y, hitTestDetails, *showDefaultMenu);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderCustomDrawExlvwa(LPDISPATCH column, ExLVwLibA::CustomDrawStageConstants drawStage, ExLVwLibA::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle, ExLVwLibA::CustomDrawReturnValuesConstants* furtherProcessing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderCustomDraw: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderCustomDraw: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", drawStage=0x%X, columnState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i), furtherProcessing=0x%X"), drawStage, columnState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom, *furtherProcessing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDividerDblClickExlvwa(LPDISPATCH column, VARIANT_BOOL* autoSizeColumn)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderDividerDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderDividerDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", autoSizeColumn=%i"), *autoSizeColumn);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDragMouseMoveExlvwa(LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pDropTargt = *dropTarget;
	if(pDropTargt) {
		BSTR text = pDropTargt->GetCaption();
		str += TEXT("ExLVwA_HeaderDragMouseMove: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderDragMouseMove: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, xListView, yListView, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderDropExlvwa(LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pDropTargt = dropTarget;
	if(pDropTargt) {
		BSTR text = pDropTargt->GetCaption();
		str += TEXT("ExLVwA_HeaderDrop: dropTarget=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderDrop: dropTarget=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderGotFocusExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_HeaderGotFocus")));
}

void __stdcall CMainDlg::HeaderItemGetDisplayInfoExlvwa(LPDISPATCH column, long requestedInfo, long* IconIndex, long maxColumnCaptionLength, BSTR* columnCaption, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		CAtlString tmp;
		tmp.Format(TEXT("ExLVwA_HeaderItemGetDisplayInfo: column=%i"), pColumn->GetIndex());
		str += tmp;
	} else {
		str += TEXT("ExLVwA_HeaderItemGetDisplayInfo: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, IconIndex=%i, maxColumnCaptionLength=%i, columnCaption=%s, dontAskAgain=%i"), requestedInfo, *IconIndex, maxColumnCaptionLength, OLE2W(*columnCaption), *dontAskAgain);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderKeyDownExlvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_HeaderKeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderKeyPressExlvwa(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_HeaderKeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderKeyUpExlvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_HeaderKeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderLostFocusExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_HeaderLostFocus")));
}

void __stdcall CMainDlg::HeaderMClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseDownExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseDown: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseDown: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseEnterExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseEnter: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseEnter: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseHoverExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseHover: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseHover: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseLeaveExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseLeave: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseLeave: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseMoveExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseMove: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseMove: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseUpExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseUp: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseUp: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderMouseWheelExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderMouseWheel: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderMouseWheel: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLECompleteDragExlvwa(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragDropExlvwa(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibA::IListViewColumn> pDropTarget = dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i"), button, shift, x, y, xListView, yListView, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
		pDraggedPicture = pData->GetData(CF_DIB, -1, 1);
		_variant_t bk;
		bk.vt = VT_DISPATCH;
		if(pDraggedPicture) {
			pDraggedPicture->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&bk.pdispVal));
		} else {
			bk.pdispVal = NULL;
		}
		controls.exlvwA->PutBkImage(bk);
	}
	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.exlvwA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::HeaderOLEDragEnterExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibA::IListViewColumn> pDropTarget = *dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, xListView, yListView, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragEnterPotentialTargetExlvwa(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_HeaderOLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragLeaveExlvwa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<ExLVwLibA::IListViewColumn> pDropTarget = dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i"), button, shift, x, y, xListView, yListView, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEDragLeavePotentialTargetExlvwa(void)
{
	AddLogEntry(TEXT("ExLVwA_HeaderOLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::HeaderOLEDragMouseMoveExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long xListView, long yListView, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibA::IListViewColumn> pDropTarget = *dropTarget;
	if(pDropTarget) {
		BSTR text = pDropTarget->GetCaption();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, xListView=%i, yListView=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, xListView, yListView, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEGiveFeedbackExlvwa(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_HeaderOLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEQueryContinueDragExlvwa(BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_HeaderOLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEReceivedNewDataExlvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLESetDataExlvwa(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, Index=%i, dataOrViewAspect=%i"), formatID, Index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOLEStartDragExlvwa(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_HeaderOLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_HeaderOLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderOwnerDrawItemExlvwa(LPDISPATCH column, ExLVwLibA::CustomDrawItemStateConstants columnState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderOwnerDrawItem: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderOwnerDrawItem: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", columnState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), columnState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderRClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderRClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderRClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderRDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderRDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderRDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderXClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderXClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderXClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HeaderXDblClickExlvwa(LPDISPATCH column, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_HeaderXDblClick: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HeaderXDblClick: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::HotItemChangedExlvwa(LPDISPATCH previousHotItem, LPDISPATCH newHotItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pPrevItem = previousHotItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExLVwA_HotItemChanged: previousHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HotItemChanged: previousHotItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewItem> pNewItem = newHotItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newHotItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::HotItemChangingExlvwa(LPDISPATCH previousHotItem, LPDISPATCH newHotItem, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pPrevItem = previousHotItem;
	if(pPrevItem) {
		BSTR text = pPrevItem->GetText();
		str += TEXT("ExLVwA_HotItemChanging: previousHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_HotItemChanging: previousHotItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewItem> pNewItem = newHotItem;
	if(pNewItem) {
		BSTR text = pNewItem->GetText();
		str += TEXT(", newHotItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", newHotItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelChange=%i"), *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::IncrementalSearchingExlvwa(BSTR currentSearchString, LONG* itemToSelect)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_IncrementalSearching: currentSearchString=%s, itemToSelect=%i"), OLE2W(currentSearchString), *itemToSelect);

	AddLogEntry(str);
}

void __stdcall CMainDlg::IncrementalSearchStringChangingExlvwa(BSTR currentSearchString, short keyCodeOfCharToBeAdded, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_IncrementalSearchStringChanging: currentSearchString=%s, keyCodeOfCharToBeAdded=%i, cancelChange=%i"), OLE2W(currentSearchString), keyCodeOfCharToBeAdded, *cancelChange);

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedColumnExlvwa(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_InsertedColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InsertedColumn: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedGroupExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_InsertedGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InsertedGroup: Group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertedItemExlvwa(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_InsertedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InsertedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingColumnExlvwa(LPDISPATCH column, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IVirtualListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_InsertingColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InsertingColumn: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingGroupExlvwa(LPDISPATCH Group, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IVirtualListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_InsertingGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InsertingGroup: Group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::InsertingItemExlvwa(LPDISPATCH listItem, VARIANT_BOOL* cancelInsertion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IVirtualListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_InsertingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InsertingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelInsertion=%i"), *cancelInsertion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::InvokeVerbFromSubItemControlExlvwa(LPDISPATCH listItem, BSTR verb)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_InvokeVerbFromSubItemControl: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_InvokeVerbFromSubItemControl: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", verb=%s"), verb);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemActivateExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemActivate: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemActivate: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", shift=%i, x=%i, y=%i"), shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemAsynchronousDrawFailedExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, ExLVwLibA::FAILEDIMAGEDETAILS* imageDetails, ExLVwLibA::ImageDrawingFailureReasonConstants failureReason, ExLVwLibA::FailedAsyncDrawReturnValuesConstants* furtherProcessing, LONG* newImageToDraw)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemAsynchronousDrawFailed: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemAsynchronousDrawFailed: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", failureReason=%i, furtherProcessing=%i, newImageToDraw=%i\n"), failureReason, *furtherProcessing, *newImageToDraw);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hImageList=0x%X\n"), imageDetails->hImageList);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.hDC=0x%X\n"), imageDetails->hDC);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.IconIndex=%i\n"), imageDetails->IconIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.OverlayIndex=%i\n"), imageDetails->OverlayIndex);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingStyle=%i\n"), imageDetails->DrawingStyle);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.DrawingEffects=%i\n"), imageDetails->DrawingEffects);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.BackColor=0x%X\n"), imageDetails->BackColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.ForeColor=0x%X\n"), imageDetails->ForeColor);
	str += tmp;
	tmp.Format(TEXT("     imageDetails.EffectColor=0x%X\n"), imageDetails->EffectColor);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemBeginDragExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemBeginDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemBeginDrag: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	CComPtr<ExLVwLibA::IListViewItems> pItems = controls.exlvwA->GetListItems();
	pItems->PutFilterType(ExLVwLibA::fpSelected, ExLVwLibA::ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(true));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
	pItems->PutFilter(ExLVwLibA::fpSelected, filter);

	ptStartDrag.x = x;
	ptStartDrag.y = y;
	controls.exlvwA->BeginDrag(controls.exlvwA->CreateItemContainer(_variant_t(pItems)), static_cast<OLE_HANDLE>(-1), NULL, NULL);
}

void __stdcall CMainDlg::ItemBeginRDragExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemBeginRDrag: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemBeginRDrag: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetDisplayInfoExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* Indent, long* groupID, LPSAFEARRAY* TileViewColumns, long maxItemTextLength, BSTR* itemText, long* OverlayIndex, long* StateImageIndex, long* itemStates, VARIANT_BOOL* dontAskAgain)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		CAtlString tmp;
		tmp.Format(TEXT("ExLVwA_ItemGetDisplayInfo: listItem=%i"), pItem->GetIndex());
		str += tmp;
	} else {
		str += TEXT("ExLVwA_ItemGetDisplayInfo: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		CAtlString tmp;
		tmp.Format(TEXT(", listItem=%i"), pSubItem->GetIndex());
		str += tmp;
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", requestedInfo=%i, IconIndex=%i, Indent=%i, groupID=%i, maxItemTextLength=%i, itemText=%s, OverlayIndex=%i, StateImageIndex=%i, itemStates=0x%X, dontAskAgain=%i"), requestedInfo, *IconIndex, *Indent, *groupID, maxItemTextLength, OLE2W(*itemText), *OverlayIndex, *StateImageIndex, *itemStates, *dontAskAgain);
	str += tmp;

	AddLogEntry(str);

	if(requestedInfo & ExLVwLibA::riItemText) {
		if(pSubItem && pSubItem->GetIndex() != 0) {
			CAtlString str;
			str.Format(TEXT("Item %i, SubItem %i"), pItem->GetIndex(), pSubItem->GetIndex());
			*itemText = _bstr_t(str).Detach();
		} else {
			CAtlString str;
			str.Format(TEXT("Item %i"), pItem->GetIndex());
			*itemText = _bstr_t(str).Detach();
		}
	}
	if(requestedInfo & ExLVwLibA::riIconIndex) {
		*IconIndex = 0;
	}
	if(requestedInfo & ExLVwLibA::riTileViewColumns) {
		CComPtr<IRecordInfo> pRecordInfo = NULL;
		CLSID clsidTILEVIEWSUBITEM = {0};
		CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
		ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, 1, 7, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
		*TileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, 2, pRecordInfo);

		ExLVwLibA::TILEVIEWSUBITEM element = {0};
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

void __stdcall CMainDlg::ItemGetGroupExlvwa(LONG itemIndex, LONG occurrenceIndex, LONG* groupIndex)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_ItemGetGroup: itemIndex=%i, occurrenceIndex=%i, groupIndex=%i"), itemIndex, occurrenceIndex, *groupIndex);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemGetInfoTipTextExlvwa(LPDISPATCH listItem, long maxInfoTipLength, BSTR* infoTipText, VARIANT_BOOL* abortToolTip)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemGetInfoTipText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemGetInfoTipText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", maxInfoTipLength=%i, infoTipText=%s, abortToolTip=%i"), maxInfoTipLength, OLE2W(*infoTipText), *abortToolTip);
	str += tmp;

	AddLogEntry(str);

	if(controls.exlvwA->GetVirtualMode() == VARIANT_FALSE) {
		if(wcslen(OLE2W(*infoTipText)) == 0) {
			CAtlString tmp;
			tmp.Format(TEXT("ID: %i\r\nItemData: 0x%X"), pItem->GetID(), pItem->GetItemData());

			*infoTipText = _bstr_t(tmp).Detach();
		} else {
			CAtlString tmp;
			tmp.Format(TEXT("%s\r\nID: %i\r\nItemData: 0x%X"), OLE2W(*infoTipText), pItem->GetID(), pItem->GetItemData());

			*infoTipText = _bstr_t(tmp).Detach();
		}
	} else {
		*infoTipText = SysAllocString(OLESTR("Hello world!"));
	}
}

void __stdcall CMainDlg::ItemGetOccurrencesCountExlvwa(LONG itemIndex, LONG* occurrencesCount)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_ItemGetOccurrencesCount: itemIndex=%i, occurrenceIndex=%i"), itemIndex, *occurrencesCount);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseEnterExlvwa(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemMouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemMouseEnter: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemMouseLeaveExlvwa(LPDISPATCH listItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemMouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemMouseLeave: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSelectionChangedExlvwa(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemSelectionChanged: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemSelectionChanged: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemSetTextExlvwa(LPDISPATCH listItem, BSTR itemText)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemSetText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemSetText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemText=%s"), OLE2W(itemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangedExlvwa(LPDISPATCH listItem, long previousStateImageIndex, long newStateImageIndex, long causedBy)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemStateImageChanged: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemStateImageChanged: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i"), previousStateImageIndex, newStateImageIndex, causedBy);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ItemStateImageChangingExlvwa(LPDISPATCH listItem, long previousStateImageIndex, long* newStateImageIndex, long causedBy, VARIANT_BOOL* cancelChange)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_ItemStateImageChanging: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ItemStateImageChanging: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousStateImageIndex=%i, newStateImageIndex=%i, causedBy=%i, cancelChange=%i"), previousStateImageIndex, *newStateImageIndex, causedBy, *cancelChange);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyDownExlvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_KeyDown: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);

	if(*keyCode == VK_F2) {
		controls.exlvwA->GetCaretItem(VARIANT_FALSE)->StartLabelEditing();
	} else if(*keyCode == VK_ESCAPE) {
		if(controls.exlvwA->GetDraggedItems()) {
			controls.exlvwA->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::KeyPressExlvwa(short* keyAscii)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_KeyPress: keyAscii=%i"), *keyAscii);

	AddLogEntry(str);
}

void __stdcall CMainDlg::KeyUpExlvwa(short* keyCode, short shift)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_KeyUp: keyCode=%i, shift=%i"), *keyCode, shift);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MapGroupWideToTotalItemIndexExlvwa(LONG groupIndex, LONG groupWideItemIndex, LONG* totalItemIndex)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_MapGroupWideToTotalItemIndex: groupIndex=%i, groupWideItemIndex=%i, totalItemIndex=%i"), groupIndex, groupWideItemIndex, *totalItemIndex);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MDblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseDown: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseDown: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseEnter: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseEnter: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseHover: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseHover: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseLeave: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseLeave: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseMove: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseMove: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseUp: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseUp: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(button == 1/*vbLeftButton*/) {
		if(controls.exlvwA->GetDraggedItems()) {
			// Are we within the client area?
			if(((ExLVwLibA::htAbove | ExLVwLibA::htBelow | ExLVwLibA::htToLeft | ExLVwLibA::htToRight) & hitTestDetails) == 0) {
				// yes
				controls.exlvwA->EndDrag(VARIANT_FALSE);
			} else {
				// no
				controls.exlvwA->EndDrag(VARIANT_TRUE);
			}
		}
	}
}

void __stdcall CMainDlg::MouseWheelExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails, long scrollAxis, short wheelDelta)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_MouseWheel: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_MouseWheel: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, scrollAxis=%i, wheelDelta=%i"), button, shift, x, y, hitTestDetails, scrollAxis, wheelDelta);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLECompleteDragExlvwa(LPDISPATCH data, long performedEffect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLECompleteDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLECompleteDrag: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", performedEffect=%i"), performedEffect);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			MessageBox(str, TEXT("Dragged files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragDropExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibA::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
		pDraggedPicture = pData->GetData(CF_DIB, -1, 1);
		_variant_t bk;
		bk.vt = VT_DISPATCH;
		if(pDraggedPicture) {
			pDraggedPicture->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&bk.pdispVal));
		} else {
			bk.pdispVal = NULL;
		}
		controls.exlvwA->PutBkImage(bk);
	}
	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.exlvwA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibA::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragEnterPotentialTargetExlvwa(long hWndPotentialTarget)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x%X"), hWndPotentialTarget);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveExlvwa(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeavePotentialTargetExlvwa(void)
{
	AddLogEntry(TEXT("ExLVwA_OLEDragLeavePotentialTarget"));
}

void __stdcall CMainDlg::OLEDragMouseMoveExlvwa(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short button, short shift, long x, long y, long hitTestDetails, long* autoHScrollVelocity, long* autoVScrollVelocity)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<ExLVwLibA::IListViewItem> pItem = *dropTarget;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i, autoHScrollVelocity=%i, autoVScrollVelocity=%i"), button, shift, x, y, hitTestDetails, *autoHScrollVelocity, *autoVScrollVelocity);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEGiveFeedbackExlvwa(long effect, VARIANT_BOOL* useDefaultCursors)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_OLEGiveFeedback: effect=%i, useDefaultCursors=%i"), effect, *useDefaultCursors);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEQueryContinueDragExlvwa(BOOL pressedEscape, short button, short shift, long* actionToContinueWith)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_OLEQueryContinueDrag: pressedEscape=%i, button=%i, shift=%i, actionToContinueWith=%i"), pressedEscape, button, shift, *actionToContinueWith);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEReceivedNewDataExlvwa(LPDISPATCH data, long formatID, long index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLEReceivedNewData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLEReceivedNewData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, index=%i, dataOrViewAspect=%i"), formatID, index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLESetDataExlvwa(LPDISPATCH data, long formatID, long Index, long dataOrViewAspect)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLESetData: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLESetData: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", formatID=%i, Index=%i, dataOrViewAspect=%i"), formatID, Index, dataOrViewAspect);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEStartDragExlvwa(LPDISPATCH data)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ExLVwA_OLEStartDrag: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ExLVwA_OLEStartDrag: data=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::OwnerDrawItemExlvwa(LPDISPATCH listItem, ExLVwLibA::OwnerDrawItemStateConstants itemState, long hDC, ExLVwLibA::RECTANGLE* drawingRectangle)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_OwnerDrawItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_OwnerDrawItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", itemState=0x%X, hDC=0x%X, drawingRectangle=(%i, %i)-(%i, %i)"), itemState, hDC, drawingRectangle->Left, drawingRectangle->Top, drawingRectangle->Right, drawingRectangle->Bottom);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_RClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_RDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RDblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwa(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ExLVwA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	if(IsComctl32Version600OrNewer()) {
		InsertGroupsA();
	}
	InsertItemsA();
	InsertFooterItemsA();
}

void __stdcall CMainDlg::RemovedColumnExlvwa(LPDISPATCH column)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IVirtualListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_RemovedColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RemovedColumn: column=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovedGroupExlvwa(LPDISPATCH Group)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IVirtualListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_RemovedGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RemovedGroup: Group=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovedItemExlvwa(LPDISPATCH listItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IVirtualListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_RemovedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RemovedItem: listItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingColumnExlvwa(LPDISPATCH column, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_RemovingColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RemovingColumn: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingGroupExlvwa(LPDISPATCH Group, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewGroup> pGroup = Group;
	if(pGroup) {
		BSTR text = pGroup->GetText(ExLVwLibA::gcHeader).Detach();
		str += TEXT("ExLVwA_RemovingGroup: Group=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RemovingGroup: Group=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RemovingItemExlvwa(LPDISPATCH listItem, VARIANT_BOOL* cancelDeletion)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_RemovingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RemovingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelDeletion=%i"), *cancelDeletion);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamedItemExlvwa(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_RenamedItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RenamedItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s"), OLE2W(previousItemText), OLE2W(newItemText));
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RenamingItemExlvwa(LPDISPATCH listItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* cancelRenaming)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_RenamingItem: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_RenamingItem: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", previousItemText=%s, newItemText=%s, cancelRenaming=%i"), OLE2W(previousItemText), OLE2W(newItemText), *cancelRenaming);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowExlvwa()
{
	AddLogEntry(CAtlString(TEXT("ExLVwA_ResizedControlWindow")));
}

void __stdcall CMainDlg::ResizingColumnExlvwa(LPDISPATCH column, long* newColumnWidth, VARIANT_BOOL* abortResizing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewColumn> pColumn = column;
	if(pColumn) {
		BSTR text = pColumn->GetCaption();
		str += TEXT("ExLVwA_ResizingColumn: column=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_ResizingColumn: column=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", newColumnWidth=%i, abortResizing=%i"), *newColumnWidth, *abortResizing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SelectedItemRangeExlvwa(LPDISPATCH firstItem, LPDISPATCH lastItem)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pFirstItem = firstItem;
	if(pFirstItem) {
		BSTR text = pFirstItem->GetText();
		str += TEXT("ExLVwA_SelectedItemRange: firstItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_SelectedItemRange: firstItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewItem> pLastItem = lastItem;
	if(pLastItem) {
		BSTR text = pLastItem->GetText();
		str += TEXT(", lastItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", lastItem=NULL");
	}

	AddLogEntry(str);
}

void __stdcall CMainDlg::SettingItemInfoTipTextExlvwa(LPDISPATCH listItem, BSTR* infoTipText, VARIANT_BOOL* abortInfoTip)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_SettingItemInfoTipText: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_SettingItemInfoTipText: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", infoTipText=%s, abortInfoTip=%i"), *infoTipText, *abortInfoTip);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::StartingLabelEditingExlvwa(LPDISPATCH listItem, VARIANT_BOOL* cancelEditing)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_StartingLabelEditing: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_StartingLabelEditing: listItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", cancelEditing=%i"), *cancelEditing);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SubItemMouseEnterExlvwa(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT("ExLVwA_SubItemMouseEnter: listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_SubItemMouseEnter: listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::SubItemMouseLeaveExlvwa(LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT("ExLVwA_SubItemMouseLeave: listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_SubItemMouseLeave: listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_XClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_XClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickExlvwa(LPDISPATCH listItem, LPDISPATCH listSubItem, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<ExLVwLibA::IListViewItem> pItem = listItem;
	if(pItem) {
		BSTR text = pItem->GetText();
		str += TEXT("ExLVwA_XDblClick: listItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("ExLVwA_XDblClick: listItem=NULL");
	}
	CComQIPtr<ExLVwLibA::IListViewSubItem> pSubItem = listSubItem;
	if(pSubItem) {
		BSTR text = pSubItem->GetText();
		str += TEXT(", listSubItem=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT(", listSubItem=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}
