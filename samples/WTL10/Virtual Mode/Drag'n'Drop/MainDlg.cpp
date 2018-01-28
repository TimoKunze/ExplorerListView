// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


const int CMainDlg::cItems = 2000;

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

	controls.oleDDCheck = GetDlgItem(IDC_OLEDDCHECK);
	controls.oleDDCheck.SetCheck(BST_CHECKED);
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

LRESULT CMainDlg::OnViewAutoArrange(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(GetMenuState(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND) & MF_CHECKED) {
		controls.exlvwU->PutAutoArrangeItems(aaiNone);
	} else {
		controls.exlvwU->PutAutoArrangeItems(aaiNormal);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewSnapToGrid(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(GetMenuState(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND) & MF_CHECKED) {
		controls.exlvwU->PutSnapToGrid(VARIANT_FALSE);
	} else {
		controls.exlvwU->PutSnapToGrid(VARIANT_TRUE);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnViewAeroDragImages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	HMENU hMainMenu = GetMenu();
	HMENU hMenu = GetSubMenu(hMainMenu, 0);
	if(GetMenuState(hMenu, ID_VIEW_AERODRAGIMAGES, MF_BYCOMMAND) & MF_CHECKED) {
		controls.exlvwU->PutOLEDragImageStyle(odistClassic);
	} else {
		controls.exlvwU->PutOLEDragImageStyle(odistAeroIfAvailable);
	}
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnAlignTop(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutItemAlignment(iaTop);
	UpdateMenu();
	return 0;
}

LRESULT CMainDlg::OnAlignLeft(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.exlvwU->PutItemAlignment(iaLeft);
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
	pColumns->Add(OLESTR("Column 3"), -1, 160, 0, alRight, 3, 1, VARIANT_TRUE, VARIANT_FALSE, VARIANT_FALSE);
}

void CMainDlg::InsertItems(void)
{
	controls.exlvwU->PuthImageList(ilSmall, HandleToLong(controls.smallImageList.m_hImageList));
	controls.exlvwU->PuthImageList(ilLarge, HandleToLong(controls.largeImageList.m_hImageList));
	controls.exlvwU->PuthImageList(ilExtraLarge, HandleToLong(controls.extralargeImageList.m_hImageList));

	controls.exlvwU->PutVirtualItemCount(VARIANT_FALSE, VARIANT_TRUE, cItems);
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
	if(controls.exlvwU->GetAutoArrangeItems() == aaiNormal) {
		CheckMenuItem(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND | MF_CHECKED);
	} else {
		CheckMenuItem(hMenu, ID_VIEW_AUTOARRANGE, MF_BYCOMMAND | MF_UNCHECKED);
	}
	if(controls.exlvwU->GetSnapToGrid() == VARIANT_FALSE) {
		CheckMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | MF_UNCHECKED);
	} else {
		CheckMenuItem(hMenu, ID_VIEW_SNAPTOGRID, MF_BYCOMMAND | MF_CHECKED);
	}
	if(controls.exlvwU->GetOLEDragImageStyle() == odistClassic) {
		CheckMenuItem(hMenu, ID_VIEW_AERODRAGIMAGES, MF_BYCOMMAND | MF_UNCHECKED);
	} else {
		CheckMenuItem(hMenu, ID_VIEW_AERODRAGIMAGES, MF_BYCOMMAND | MF_CHECKED);
	}

	hMenu = GetSubMenu(hMainMenu, 1);
	switch(controls.exlvwU->GetItemAlignment()) {
		case iaTop:
			CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_TOP, MF_BYCOMMAND);
			break;
		case iaLeft:
			CheckMenuRadioItem(hMenu, ID_ALIGN_TOP, ID_ALIGN_LEFT, ID_ALIGN_LEFT, MF_BYCOMMAND);
			break;
	}

	if(controls.exlvwU->GetAutoArrangeItems() == VARIANT_FALSE) {
		freeFloat = TRUE;
	} else {
		freeFloat = FALSE;
	}
}

void __stdcall CMainDlg::AbortedDragExlvwu()
{
	controls.exlvwU->PutRefDropHilitedItem(NULL);
	lastDropTarget = -1;
	if(IsComctl32Version600OrNewer() && !freeFloat) {
		controls.exlvwU->SetInsertMarkPosition(impNowhere, NULL);
	}
}

void __stdcall CMainDlg::CreatedHeaderControlWindowExlvwu(long /*hWndHeader*/)
{
	InsertColumns();
}

void __stdcall CMainDlg::DragMouseMoveExlvwu(LPDISPATCH* dropTarget, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	CComQIPtr<IListViewItem> pDropTarget = *dropTarget;
	int newDropTarget;
	if(!pDropTarget) {
		newDropTarget = -1;
		//controls.exlvwU->PutShowDragImage(VARIANT_FALSE);
	} else {
		newDropTarget = pDropTarget->GetIndex();
		//controls.exlvwU->PutShowDragImage(VARIANT_TRUE);
	}
	lastDropTarget = newDropTarget;

	if(IsComctl32Version600OrNewer() && !freeFloat) {
		InsertMarkPositionConstants insertMarkRelativePosition;
		CComPtr<IListViewItem> pItem = NULL;
		controls.exlvwU->GetClosestInsertMarkPosition(x, y, &insertMarkRelativePosition, &pItem);
		controls.exlvwU->SetInsertMarkPosition(insertMarkRelativePosition, pItem);
	}
}

void __stdcall CMainDlg::DropExlvwu(LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/)
{
	if(bRightDrag) {
		HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_DRAGMENU));
		HMENU hPopupMenu = GetSubMenu(hMenu, 0);

		POINT pt = {x, y};
		::ClientToScreen(static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), &pt);
		UINT selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.exlvwU->GethWnd())), NULL);
		DestroyMenu(hPopupMenu);
		DestroyMenu(hMenu);

		switch(selectedCmd) {
			case ID_ACTIONS_CANCEL:
				AbortedDragExlvwu();
				return;
				break;
			case ID_ACTIONS_COPYITEM:
				MessageBox(TEXT("TODO: Copy the item."), TEXT("Command not implemented"), MB_ICONINFORMATION);
				AbortedDragExlvwu();
				return;
				break;
			case ID_ACTIONS_MOVEITEM:
				// fall through
				break;
		}
	}

	MessageBox(TEXT("TODO: Move the item."), TEXT("Command not implemented"), MB_ICONINFORMATION);
	AbortedDragExlvwu();
}

void __stdcall CMainDlg::ItemBeginDragExlvwu(LPDISPATCH /*listItem*/, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/)
{
	bRightDrag = FALSE;
	lastDropTarget = -1;
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	pItems->PutFilterType(fpSelected, ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(true));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
	pItems->PutFilter(fpSelected, filter);

	ptStartDrag.x = x;
	ptStartDrag.y = y;
	if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
		IDataObject* p = NULL;
		controls.exlvwU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.exlvwU->CreateItemContainer(_variant_t(pItems)), -1);
	} else {
		controls.exlvwU->BeginDrag(controls.exlvwU->CreateItemContainer(_variant_t(pItems)), static_cast<OLE_HANDLE>(-1), &szHotSpotOffset.cx, &szHotSpotOffset.cy);
	}
}

void __stdcall CMainDlg::ItemBeginRDragExlvwu(LPDISPATCH /*listItem*/, LPDISPATCH /*listSubItem*/, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/)
{
	bRightDrag = TRUE;
	lastDropTarget = -1;
	CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
	pItems->PutFilterType(fpSelected, ftIncluding);
	_variant_t filter;
	filter.Clear();
	CComSafeArray<VARIANT> arr;
	arr.Create(1, 1);
	arr.SetAt(1, _variant_t(true));
	filter.parray = arr.Detach();
	filter.vt = VT_ARRAY | VT_VARIANT;     // NOTE: ExplorerListView expects an array of VARIANTs!
	pItems->PutFilter(fpSelected, filter);

	ptStartDrag.x = x;
	ptStartDrag.y = y;
	if(controls.oleDDCheck.GetCheck() == BST_CHECKED) {
		IDataObject* p = NULL;
		controls.exlvwU->OLEDrag(reinterpret_cast<long*>(&p), odeCopyOrMove, static_cast<OLE_HANDLE>(-1), controls.exlvwU->CreateItemContainer(_variant_t(pItems)), -1);
	} else {
		controls.exlvwU->BeginDrag(controls.exlvwU->CreateItemContainer(_variant_t(pItems)), static_cast<OLE_HANDLE>(-1), &szHotSpotOffset.cx, &szHotSpotOffset.cy);
	}
}

void __stdcall CMainDlg::ItemGetDisplayInfoExlvwu(LPDISPATCH listItem, LPDISPATCH listSubItem, long requestedInfo, long* IconIndex, long* /*Indent*/, long* /*groupID*/, LPSAFEARRAY* TileViewColumns, long /*maxItemTextLength*/, BSTR* itemText, long* /*OverlayIndex*/, long* /*StateImageIndex*/, long* /*itemStates*/, VARIANT_BOOL* /*dontAskAgain*/)
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
		ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, 1, 0, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
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

void __stdcall CMainDlg::KeyDownExlvwu(short* keyCode, short /*shift*/)
{
	if(*keyCode == VK_ESCAPE) {
		if(controls.exlvwU->GetDraggedItems()) {
			controls.exlvwU->EndDrag(VARIANT_TRUE);
		}
	}
}

void __stdcall CMainDlg::MouseUpExlvwu(LPDISPATCH /*listItem*/, LPDISPATCH /*listSubItem*/, short button, short /*shift*/, long /*x*/, long /*y*/, long hitTestDetails)
{
	if(controls.oleDDCheck.GetCheck() == BST_UNCHECKED) {
		if(((button == 1) && !bRightDrag) || ((button == 2) && bRightDrag)) {
			if(controls.exlvwU->GetDraggedItems()) {
				// Are we within the client area?
				if(((htAbove | htBelow | htToLeft | htToRight) & hitTestDetails) == 0) {
					// yes
					controls.exlvwU->EndDrag(VARIANT_FALSE);
				} else {
					// no
					controls.exlvwU->EndDrag(VARIANT_TRUE);
				}
			}
		}
	}
}

void __stdcall CMainDlg::OLECompleteDragExlvwu(LPDISPATCH data, long /*performedEffect*/)
{
	CComQIPtr<IDataObject> pData = data;
	if(pData) {
		FORMATETC format = {CF_TARGETCLSID, NULL, 1, -1, TYMED_HGLOBAL};
		if(pData->QueryGetData(&format) == S_OK) {
			STGMEDIUM storageMedium = {0};
			if(pData->GetData(&format, &storageMedium) == S_OK) {
				LPGUID pCLSIDOfTarget = reinterpret_cast<LPGUID>(GlobalLock(storageMedium.hGlobal));
				if(*pCLSIDOfTarget == CLSID_RecycleBin) {
					// remove the dragged items
					MessageBox(TEXT("TODO: Remove the item."), TEXT("Command not implemented"), MB_ICONINFORMATION);
				}
				GlobalUnlock(storageMedium.hGlobal);
				ReleaseStgMedium(&storageMedium);
			}
		}
	}
}

void __stdcall CMainDlg::OLEDragDropExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* /*dropTarget*/, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
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
		} else {
			CComBSTR str = NULL;
			if(pData->GetFormat(CF_UNICODETEXT, -1, 1) != VARIANT_FALSE) {
				str = pData->GetData(CF_UNICODETEXT, -1, 1).bstrVal;
			} else if(pData->GetFormat(CF_TEXT, -1, 1) != VARIANT_FALSE) {
				str = pData->GetData(CF_TEXT, -1, 1).bstrVal;
			} else if(pData->GetFormat(CF_OEMTEXT, -1, 1) != VARIANT_FALSE) {
				str = pData->GetData(CF_OEMTEXT, -1, 1).bstrVal;
			}

			if(str != TEXT("")) {
				// insert a new item
				CAtlString tmp;
				tmp.Format(TEXT("TODO: Insert a new item named %s."), str);
				MessageBox(tmp, TEXT("Command not implemented"), MB_ICONINFORMATION);
			} else if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
				// insert a new item for each file/folder
				_variant_t v = pData->GetData(CF_HDROP, -1, 1);
				CComSafeArray<BSTR> arr;
				arr.Attach(v.parray);

				BOOL hasPositions = FALSE;
				LPPOINT pPositions = NULL;
				SIZE spacingLarge = {0};
				SIZE spacingSmall = {0};
				STGMEDIUM storageMedium = {0};
				CComQIPtr<IDataObject> pNativeData = data;
				if(pNativeData) {
					FORMATETC format = {CF_SHELLIDLISTOFFSET, NULL, 1, -1, TYMED_HGLOBAL};
					if(pNativeData->QueryGetData(&format) == S_OK) {
						if(pNativeData->GetData(&format, &storageMedium) == S_OK) {
							hasPositions = TRUE;
							pPositions = reinterpret_cast<LPPOINT>(GlobalLock(storageMedium.hGlobal));

							/* The first POINTAPI structure specifies the position (in client coordinates of the source
							   window) of the mouse cursor within the drag image. We don't need it. */
							pPositions++;
							/* the remainder of the POINTAPI structures specify the locations of the individual files relative to
							   this position */

							if(controls.exlvwU->GetView() == vIcons) {
								spacingLarge.cx = 1;
								spacingLarge.cy = 1;
								spacingSmall.cx = 1;
								spacingSmall.cy = 1;
							} else if(controls.exlvwU->GetView() == vTiles) {
								controls.exlvwU->GetIconSpacing(vIcons, &spacingLarge.cx, &spacingLarge.cy);
								controls.exlvwU->GetIconSpacing(vSmallIcons, &spacingSmall.cx, &spacingSmall.cy);
							}
							POINT origin = {0};
							controls.exlvwU->GetOrigin(&origin.x, &origin.y);
							x += origin.x;
							y += origin.y;
						}
					}
				}

				CComPtr<IListViewItems> pItems = controls.exlvwU->GetListItems();
				for(LONG i = arr.GetLowerBound(); i <= arr.GetUpperBound(); i++) {
					CAtlString tmp;
					tmp.Format(TEXT("TODO: Insert a new item named %s."), CAtlString(arr.GetAt(i)));
					MessageBox(tmp, TEXT("Command not implemented"), MB_ICONINFORMATION);
				}
				if(pPositions) {
					GlobalUnlock(storageMedium.hGlobal);
					ReleaseStgMedium(&storageMedium);
				}
				arr.Detach();
			}
		}
	}

	lastDropTarget = -1;
	if(IsComctl32Version600OrNewer() && !freeFloat) {
		controls.exlvwU->SetInsertMarkPosition(impNowhere, NULL);
	}

	if(shift & 2/*vbCtrlMask*/) {
		*effect = odeCopy;
	}
	if(shift & 4/*vbAltMask*/) {
		*effect = odeLink;
	}
}

void __stdcall CMainDlg::OLEDragEnterExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	_bstr_t dropActionDescription;
	DropDescriptionIconConstants dropDescriptionIcon = ddiNoDrop;
	_bstr_t dropTargetDescription;

	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
			*effect = odeCopy;
			dropActionDescription = TEXT("Use as background image");
			dropDescriptionIcon = ddiCopy;
		} else {
			CComQIPtr<IListViewItem> pDropTarget = *dropTarget;
			long newDropTarget;
			if(!pDropTarget) {
				newDropTarget = -1;
			} else {
				newDropTarget = pDropTarget->GetIndex();
			}
			lastDropTarget = newDropTarget;

			CComPtr<IListViewItem> pItem = NULL;
			InsertMarkPositionConstants insertMarkRelativePosition = impNowhere;
			if(IsComctl32Version600OrNewer() && !freeFloat) {
				controls.exlvwU->GetClosestInsertMarkPosition(x, y, &insertMarkRelativePosition, &pItem);
				controls.exlvwU->SetInsertMarkPosition(insertMarkRelativePosition, pItem);
				if(pItem) {
					dropTargetDescription = pItem->GetText();
				}
			}

			*effect = odeCopy;
			if(shift & 1/*vbShiftMask*/) {
				*effect = odeMove;
			}
			if(shift & 2/*vbCtrlMask*/) {
				*effect = odeCopy;
			}
			if(shift & 4/*vbAltMask*/) {
				*effect = odeLink;
			}

			if(!pItem) {
				if(freeFloat) {
					switch(*effect) {
						case odeMove:
							dropActionDescription = TEXT("Move here");
							dropDescriptionIcon = ddiMove;
							break;
						case odeLink:
							dropActionDescription = TEXT("Create link here");
							dropDescriptionIcon = ddiLink;
							break;
						default:
							dropActionDescription = TEXT("Insert copy here");
							dropDescriptionIcon = ddiCopy;
							break;
					}
				} else {
					dropDescriptionIcon = ddiNoDrop;
					dropActionDescription = TEXT("Cannot place here");
				}
			} else {
				switch(*effect) {
					case odeMove:
						switch(insertMarkRelativePosition) {
							case impAfter:
								dropActionDescription = TEXT("Move after %1");
								break;
							case impBefore:
								dropActionDescription = TEXT("Move before %1");
								break;
							default:
								dropActionDescription = TEXT("Move to %1");
								break;
						}
						dropDescriptionIcon = ddiMove;
						break;
					case odeLink:
						switch(insertMarkRelativePosition) {
							case impAfter:
								dropActionDescription = TEXT("Create link after %1");
								break;
							case impBefore:
								dropActionDescription = TEXT("Create link before %1");
								break;
							default:
								dropActionDescription = TEXT("Create link in %1");
								break;
						}
						dropDescriptionIcon = ddiLink;
						break;
					default:
						switch(insertMarkRelativePosition) {
							case impAfter:
								dropActionDescription = TEXT("Insert copy after %1");
								break;
							case impBefore:
								dropActionDescription = TEXT("Insert copy before %1");
								break;
							default:
								dropActionDescription = TEXT("Copy to %1");
								break;
						}
						dropDescriptionIcon = ddiCopy;
						break;
				}
			}
		}
		pData->SetDropDescription(dropTargetDescription, dropActionDescription, dropDescriptionIcon);
	}
}

void __stdcall CMainDlg::OLEDragLeaveExlvwu(LPDISPATCH /*data*/, LPDISPATCH /*dropTarget*/, short /*button*/, short /*shift*/, long /*x*/, long /*y*/, long /*hitTestDetails*/)
{
	lastDropTarget = -1;
	if(IsComctl32Version600OrNewer() && !freeFloat) {
		controls.exlvwU->SetInsertMarkPosition(impNowhere, NULL);
	}
}

void __stdcall CMainDlg::OLEDragMouseMoveExlvwu(LPDISPATCH data, long* effect, LPDISPATCH* dropTarget, short /*button*/, short shift, long x, long y, long /*hitTestDetails*/, long* /*autoHScrollVelocity*/, long* /*autoVScrollVelocity*/)
{
	_bstr_t dropActionDescription;
	DropDescriptionIconConstants dropDescriptionIcon = ddiNoDrop;
	_bstr_t dropTargetDescription;

	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		if(pData->GetFormat(CF_DIB, -1, 1) != VARIANT_FALSE) {
			*effect = odeCopy;
			dropActionDescription = TEXT("Use as background image");
			dropDescriptionIcon = ddiCopy;
		} else {
			CComQIPtr<IListViewItem> pDropTarget = *dropTarget;
			long newDropTarget;
			if(!pDropTarget) {
				newDropTarget = -1;
			} else {
				newDropTarget = pDropTarget->GetIndex();
			}
			lastDropTarget = newDropTarget;

			CComPtr<IListViewItem> pItem = NULL;
			InsertMarkPositionConstants insertMarkRelativePosition = impNowhere;
			if(IsComctl32Version600OrNewer() && !freeFloat) {
				controls.exlvwU->GetClosestInsertMarkPosition(x, y, &insertMarkRelativePosition, &pItem);
				controls.exlvwU->SetInsertMarkPosition(insertMarkRelativePosition, pItem);
				if(pItem) {
					dropTargetDescription = pItem->GetText();
				}
			}

			*effect = odeCopy;
			if(shift & 1/*vbShiftMask*/) {
				*effect = odeMove;
			}
			if(shift & 2/*vbCtrlMask*/) {
				*effect = odeCopy;
			}
			if(shift & 4/*vbAltMask*/) {
				*effect = odeLink;
			}

			if(!pItem) {
				if(freeFloat) {
					switch(*effect) {
						case odeMove:
							dropActionDescription = TEXT("Move here");
							dropDescriptionIcon = ddiMove;
							break;
						case odeLink:
							dropActionDescription = TEXT("Create link here");
							dropDescriptionIcon = ddiLink;
							break;
						default:
							dropActionDescription = TEXT("Insert copy here");
							dropDescriptionIcon = ddiCopy;
							break;
					}
				} else {
					dropDescriptionIcon = ddiNoDrop;
					dropActionDescription = TEXT("Cannot place here");
				}
			} else {
				switch(*effect) {
					case odeMove:
						switch(insertMarkRelativePosition) {
							case impAfter:
								dropActionDescription = TEXT("Move after %1");
								break;
							case impBefore:
								dropActionDescription = TEXT("Move before %1");
								break;
							default:
								dropActionDescription = TEXT("Move to %1");
								break;
						}
						dropDescriptionIcon = ddiMove;
						break;
					case odeLink:
						switch(insertMarkRelativePosition) {
							case impAfter:
								dropActionDescription = TEXT("Create link after %1");
								break;
							case impBefore:
								dropActionDescription = TEXT("Create link before %1");
								break;
							default:
								dropActionDescription = TEXT("Create link in %1");
								break;
						}
						dropDescriptionIcon = ddiLink;
						break;
					default:
						switch(insertMarkRelativePosition) {
							case impAfter:
								dropActionDescription = TEXT("Insert copy after %1");
								break;
							case impBefore:
								dropActionDescription = TEXT("Insert copy before %1");
								break;
							default:
								dropActionDescription = TEXT("Copy to %1");
								break;
						}
						dropDescriptionIcon = ddiCopy;
						break;
				}
			}
		}
		pData->SetDropDescription(dropTargetDescription, dropActionDescription, dropDescriptionIcon);
	}
}

void __stdcall CMainDlg::OLESetDataExlvwu(LPDISPATCH data, long formatID, long /*Index*/, long /*dataOrViewAspect*/)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		switch(formatID) {
			case CF_TEXT:
			case CF_OEMTEXT:
			case CF_UNICODETEXT: {
				CComQIPtr<IEnumVARIANT> pEnum = controls.exlvwU->GetDraggedItems();
				pEnum->Reset();     // NOTE: It's important to call Reset()!
				ULONG ul = 0;
				_variant_t v;
				v.Clear();
				CComBSTR str = L"";
				while(pEnum->Next(1, &v, &ul) == S_OK) {
					CComQIPtr<IListViewItem> pItem = v.pdispVal;
					str.AppendBSTR(pItem->GetText());
					str.Append(TEXT("\r\n"));
				}
				pData->SetData(formatID, _variant_t(str), -1, 1);
				break;
			}

			case CF_BITMAP:
			case CF_DIB:
			case CF_DIBV5: {
				_variant_t v;
				v.vt = VT_DISPATCH;
				if(pDraggedPicture) {
					pDraggedPicture->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&v.pdispVal));
				} else {
					v.pdispVal = NULL;
				}
				pData->SetData(formatID, v, -1, 1);
				break;
			}

			case CF_PALETTE:
				if(pDraggedPicture) {
					OLE_HANDLE h = NULL;
					pDraggedPicture->get_hPal(&h);
					pData->SetData(formatID, _variant_t(h), -1, 1);
				}
				break;
		}
	}
}

void __stdcall CMainDlg::OLEStartDragExlvwu(LPDISPATCH data)
{
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		_variant_t invalidVariant;
		invalidVariant.vt = VT_ERROR;

		pData->SetData(CF_TEXT, invalidVariant, -1, 1);
		pData->SetData(CF_OEMTEXT, invalidVariant, -1, 1);
		pData->SetData(CF_UNICODETEXT, invalidVariant, -1, 1);
		pData->SetData(CF_HDROP, invalidVariant, -1, 1);
		if(pDraggedPicture) {
			pData->SetData(CF_PALETTE, invalidVariant, -1, 1);
			pData->SetData(CF_BITMAP, invalidVariant, -1, 1);
			pData->SetData(CF_DIB, invalidVariant, -1, 1);
			pData->SetData(CF_DIBV5, invalidVariant, -1, 1);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowExlvwu(long /*hWnd*/)
{
	InsertItems();
	UpdateMenu();
}
