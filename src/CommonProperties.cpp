// CommonProperties.cpp: The control's "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));

	controls.itemBoundingBoxList.SubclassWindow(GetDlgItem(IDC_ITEMBOUNDINGBOX));
	hStateImageList = SetupStateImageList(controls.itemBoundingBoxList.GetImageList(LVSIL_STATE));
	controls.itemBoundingBoxList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.itemBoundingBoxList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER, LVS_EX_DOUBLEBUFFER | LVS_EX_FULLROWSELECT);
	controls.itemBoundingBoxList.AddColumn(TEXT(""), 0);

	controls.callBackMaskList.SubclassWindow(GetDlgItem(IDC_CALLBACKMASKBOX));
	hStateImageList = SetupStateImageList(controls.callBackMaskList.GetImageList(LVSIL_STATE));
	controls.callBackMaskList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.callBackMaskList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER, LVS_EX_DOUBLEBUFFER | LVS_EX_FULLROWSELECT);
	controls.callBackMaskList.AddColumn(TEXT(""), 0);

	// setup the toolbar
	CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return TRUE;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(pNotificationDetails->hwndFrom != controls.disabledEventsList) {
		return 0;
	}

	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	switch(pDetails->iItem) {
		case 0:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MClick, XClick, MouseEnter, MouseHover, MouseLeave, ItemMouseEnter, ItemMouseLeave, SubItemMouseEnter, SubItemMouseLeave, MouseMove, MouseWheel"))));
			break;
		case 1:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("HeaderMouseDown, HeaderMouseUp, HeaderMouseEnter, HeaderMouseHover, HeaderMouseLeave, ColumnMouseEnter, ColumnMouseLeave, HeaderMouseMove, HeaderMouseWheel"))));
			break;
		case 2:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("EditMouseDown, EditMouseUp, EditMouseEnter, EditMouseHover, EditMouseLeave, EditMouseMove, EditMouseWheel, EditClick, EditDblClick, EditMClick, EditMDblClick, EditRClick, EditRDblClick, EditXClick, EditXDblClick"))));
			break;
		case 3:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
			break;
		case 4:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("HeaderClick, HeaderDblClick, HeaderMClick, HeaderMDblClick, HeaderRClick, HeaderRDblClick, HeaderXClick, HeaderXDblClick"))));
			break;
		case 5:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress, IncrementalSearchStringChanging, IncrementalSearching"))));
			break;
		case 6:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("HeaderKeyDown, HeaderKeyUp, HeaderKeyPress"))));
			break;
		case 7:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("EditKeyDown, EditKeyUp, EditKeyPress"))));
			break;
		case 8:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingItem, InsertedItem"))));
			break;
		case 9:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingColumn, InsertedColumn"))));
			break;
		case 10:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingGroup, InsertedGroup"))));
			break;
		case 11:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingItem, RemovedItem"))));
			break;
		case 12:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingColumn, RemovedColumn"))));
			break;
		case 13:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingGroup, RemovedGroup"))));
			break;
		case 14:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeItemData"))));
			break;
		case 15:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeColumnData"))));
			break;
		case 16:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeFooterItemData"))));
			break;
		case 17:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CustomDraw, GroupCustomDraw"))));
			break;
		case 18:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("HeaderCustomDraw"))));
			break;
		case 19:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("HotItemChanging, HotItemChanged"))));
			break;
		case 20:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("GetSubItemControl"))));
			break;
		case 21:
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("ItemGetDisplayInfo"))));
			break;
	}
	ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IExplorerListView, &IID_IExplorerListView> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IExplorerListView, &IID_IExplorerListView> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(FALSE, NULL, TEXT("ExplorerListView Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<IExplorerListView, &IID_IExplorerListView> pControl(m_ppUnk[object]);
		if(pControl) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			pControl->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(0, LVIS_STATEIMAGEMASK), l, deListMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(1, LVIS_STATEIMAGEMASK), l, deHeaderMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(2, LVIS_STATEIMAGEMASK), l, deEditMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(3, LVIS_STATEIMAGEMASK), l, deListClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(4, LVIS_STATEIMAGEMASK), l, deHeaderClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(5, LVIS_STATEIMAGEMASK), l, deListKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(6, LVIS_STATEIMAGEMASK), l, deHeaderKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(7, LVIS_STATEIMAGEMASK), l, deEditKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(8, LVIS_STATEIMAGEMASK), l, deItemInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(9, LVIS_STATEIMAGEMASK), l, deColumnInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(10, LVIS_STATEIMAGEMASK), l, deGroupInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(11, LVIS_STATEIMAGEMASK), l, deItemDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(12, LVIS_STATEIMAGEMASK), l, deColumnDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(13, LVIS_STATEIMAGEMASK), l, deGroupDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(14, LVIS_STATEIMAGEMASK), l, deFreeItemData);
			SetBit(controls.disabledEventsList.GetItemState(15, LVIS_STATEIMAGEMASK), l, deFreeColumnData);
			SetBit(controls.disabledEventsList.GetItemState(16, LVIS_STATEIMAGEMASK), l, deFreeFooterItemData);
			SetBit(controls.disabledEventsList.GetItemState(17, LVIS_STATEIMAGEMASK), l, deCustomDraw);
			SetBit(controls.disabledEventsList.GetItemState(18, LVIS_STATEIMAGEMASK), l, deHeaderCustomDraw);
			SetBit(controls.disabledEventsList.GetItemState(19, LVIS_STATEIMAGEMASK), l, deHotItemChangeEvents);
			SetBit(controls.disabledEventsList.GetItemState(20, LVIS_STATEIMAGEMASK), l, deGetSubItemControl);
			SetBit(controls.disabledEventsList.GetItemState(21, LVIS_STATEIMAGEMASK), l, deItemGetDisplayInfo);
			pControl->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

			ItemBoundingBoxDefinitionConstants boundingBox = static_cast<ItemBoundingBoxDefinitionConstants>(0);
			pControl->get_ItemBoundingBoxDefinition(&boundingBox);
			l = static_cast<LONG>(boundingBox);
			SetBit(controls.itemBoundingBoxList.GetItemState(0, LVIS_STATEIMAGEMASK), l, ibbdItemStateImage);
			SetBit(controls.itemBoundingBoxList.GetItemState(1, LVIS_STATEIMAGEMASK), l, ibbdItemIcon);
			SetBit(controls.itemBoundingBoxList.GetItemState(2, LVIS_STATEIMAGEMASK), l, ibbdItemLabel);
			pControl->put_ItemBoundingBoxDefinition(static_cast<ItemBoundingBoxDefinitionConstants>(l));

			CallBackMaskConstants callBackMask = static_cast<CallBackMaskConstants>(0);
			pControl->get_CallBackMask(&callBackMask);
			l = static_cast<LONG>(callBackMask);
			SetBit(controls.callBackMaskList.GetItemState(0, LVIS_STATEIMAGEMASK), l, cbmCaret);
			SetBit(controls.callBackMaskList.GetItemState(1, LVIS_STATEIMAGEMASK), l, cbmSelected);
			SetBit(controls.callBackMaskList.GetItemState(2, LVIS_STATEIMAGEMASK), l, cbmGhosted);
			SetBit(controls.callBackMaskList.GetItemState(3, LVIS_STATEIMAGEMASK), l, cbmDropHilited);
			SetBit(controls.callBackMaskList.GetItemState(4, LVIS_STATEIMAGEMASK), l, cbmGlowing);
			SetBit(controls.callBackMaskList.GetItemState(5, LVIS_STATEIMAGEMASK), l, cbmActivating);
			SetBit(controls.callBackMaskList.GetItemState(6, LVIS_STATEIMAGEMASK), l, cbmOverlayIndex);
			SetBit(controls.callBackMaskList.GetItemState(7, LVIS_STATEIMAGEMASK), l, cbmStateImageIndex);
			pControl->put_CallBackMask(static_cast<CallBackMaskConstants>(l));
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	DisabledEventsConstants* pDisabledEvents = reinterpret_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ItemBoundingBoxDefinitionConstants* pBoundingBox = reinterpret_cast<ItemBoundingBoxDefinitionConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(ItemBoundingBoxDefinitionConstants)));
		if(pBoundingBox) {
			CallBackMaskConstants* pCallBackMask = reinterpret_cast<CallBackMaskConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(CallBackMaskConstants)));
			if(pCallBackMask) {
				for(UINT object = 0; object < m_nObjects; ++object) {
					CComQIPtr<IExplorerListView, &IID_IExplorerListView> pControl(m_ppUnk[object]);
					if(pControl) {
						pControl->get_DisabledEvents(&pDisabledEvents[object]);
						pControl->get_ItemBoundingBoxDefinition(&pBoundingBox[object]);
						pControl->get_CallBackMask(&pCallBackMask[object]);
					}
				}

				// fill the listboxes
				LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
				properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
				controls.disabledEventsList.DeleteAllItems();
				controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events (listview)"));
				controls.disabledEventsList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, deListMouseEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(1, 0, TEXT("Mouse events (header)"));
				controls.disabledEventsList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, deHeaderMouseEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(2, 0, TEXT("Mouse events (contained edit control)"));
				controls.disabledEventsList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, deEditMouseEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(3, 0, TEXT("Click events (listview)"));
				controls.disabledEventsList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, deListClickEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(4, 0, TEXT("Click events (header)"));
				controls.disabledEventsList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, deHeaderClickEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(5, 0, TEXT("Keyboard events (listview)"));
				controls.disabledEventsList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, deListKeyboardEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(6, 0, TEXT("Keyboard events (contained header control)"));
				controls.disabledEventsList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, deHeaderKeyboardEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(7, 0, TEXT("Keyboard events (contained edit control)"));
				controls.disabledEventsList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, deEditKeyboardEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(8, 0, TEXT("Item insertions"));
				controls.disabledEventsList.SetItemState(8, CalculateStateImageMask(m_nObjects, pl, deItemInsertionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(9, 0, TEXT("Column insertions"));
				controls.disabledEventsList.SetItemState(9, CalculateStateImageMask(m_nObjects, pl, deColumnInsertionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(10, 0, TEXT("Group insertions"));
				controls.disabledEventsList.SetItemState(10, CalculateStateImageMask(m_nObjects, pl, deGroupInsertionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(11, 0, TEXT("Item deletions"));
				controls.disabledEventsList.SetItemState(11, CalculateStateImageMask(m_nObjects, pl, deItemDeletionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(12, 0, TEXT("Column deletions"));
				controls.disabledEventsList.SetItemState(12, CalculateStateImageMask(m_nObjects, pl, deColumnDeletionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(13, 0, TEXT("Group deletions"));
				controls.disabledEventsList.SetItemState(13, CalculateStateImageMask(m_nObjects, pl, deGroupDeletionEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(14, 0, TEXT("FreeItemData event"));
				controls.disabledEventsList.SetItemState(14, CalculateStateImageMask(m_nObjects, pl, deFreeItemData), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(15, 0, TEXT("FreeColumnData event"));
				controls.disabledEventsList.SetItemState(15, CalculateStateImageMask(m_nObjects, pl, deFreeColumnData), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(16, 0, TEXT("FreeFooterItemData event"));
				controls.disabledEventsList.SetItemState(16, CalculateStateImageMask(m_nObjects, pl, deFreeFooterItemData), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(17, 0, TEXT("CustomDraw + GroupCustomDraw events"));
				controls.disabledEventsList.SetItemState(17, CalculateStateImageMask(m_nObjects, pl, deCustomDraw), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(18, 0, TEXT("HeaderCustomDraw event"));
				controls.disabledEventsList.SetItemState(18, CalculateStateImageMask(m_nObjects, pl, deHeaderCustomDraw), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(19, 0, TEXT("Hot item changing events"));
				controls.disabledEventsList.SetItemState(19, CalculateStateImageMask(m_nObjects, pl, deHotItemChangeEvents), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(20, 0, TEXT("GetSubItemControl event"));
				controls.disabledEventsList.SetItemState(20, CalculateStateImageMask(m_nObjects, pl, deGetSubItemControl), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.AddItem(21, 0, TEXT("ItemGetDisplayInfo event"));
				controls.disabledEventsList.SetItemState(21, CalculateStateImageMask(m_nObjects, pl, deItemGetDisplayInfo), LVIS_STATEIMAGEMASK);
				controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				pl = reinterpret_cast<LONG*>(pBoundingBox);
				properties.hWndCheckMarksAreSetFor = controls.itemBoundingBoxList;
				controls.itemBoundingBoxList.DeleteAllItems();
				controls.itemBoundingBoxList.AddItem(0, 0, TEXT("State image"));
				controls.itemBoundingBoxList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, ibbdItemStateImage), LVIS_STATEIMAGEMASK);
				controls.itemBoundingBoxList.AddItem(1, 0, TEXT("Icon"));
				controls.itemBoundingBoxList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, ibbdItemIcon), LVIS_STATEIMAGEMASK);
				controls.itemBoundingBoxList.AddItem(2, 0, TEXT("Text label"));
				controls.itemBoundingBoxList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, ibbdItemLabel), LVIS_STATEIMAGEMASK);
				controls.itemBoundingBoxList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				pl = reinterpret_cast<LONG*>(pCallBackMask);
				properties.hWndCheckMarksAreSetFor = controls.callBackMaskList;
				controls.callBackMaskList.DeleteAllItems();
				controls.callBackMaskList.AddItem(0, 0, TEXT("Caret property"));
				controls.callBackMaskList.SetItemState(0, CalculateStateImageMask(m_nObjects, pl, cbmCaret), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(1, 0, TEXT("Selected property"));
				controls.callBackMaskList.SetItemState(1, CalculateStateImageMask(m_nObjects, pl, cbmSelected), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(2, 0, TEXT("Ghosted property"));
				controls.callBackMaskList.SetItemState(2, CalculateStateImageMask(m_nObjects, pl, cbmGhosted), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(3, 0, TEXT("DropHilited property"));
				controls.callBackMaskList.SetItemState(3, CalculateStateImageMask(m_nObjects, pl, cbmDropHilited), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(4, 0, TEXT("Glowing property"));
				controls.callBackMaskList.SetItemState(4, CalculateStateImageMask(m_nObjects, pl, cbmGlowing), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(5, 0, TEXT("Activating property"));
				controls.callBackMaskList.SetItemState(5, CalculateStateImageMask(m_nObjects, pl, cbmActivating), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(6, 0, TEXT("OverlayIndex property"));
				controls.callBackMaskList.SetItemState(6, CalculateStateImageMask(m_nObjects, pl, cbmOverlayIndex), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.AddItem(7, 0, TEXT("StateImageIndex property"));
				controls.callBackMaskList.SetItemState(7, CalculateStateImageMask(m_nObjects, pl, cbmStateImageIndex), LVIS_STATEIMAGEMASK);
				controls.callBackMaskList.SetColumnWidth(0, LVSCW_AUTOSIZE);

				properties.hWndCheckMarksAreSetFor = NULL;

				HeapFree(GetProcessHeap(), 0, pCallBackMask);
			}
			HeapFree(GetProcessHeap(), 0, pBoundingBox);
		}
		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}