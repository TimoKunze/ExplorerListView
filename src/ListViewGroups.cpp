// ListViewGroups.cpp: Manages a collection of ListViewGroup objects

#include "stdafx.h"
#include "ListViewGroups.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewGroups::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewGroups, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListViewGroups::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListViewGroups>* pLvwGroupsObj = NULL;
	CComObject<ListViewGroups>::CreateInstance(&pLvwGroupsObj);
	pLvwGroupsObj->AddRef();

	// clone all settings
	pLvwGroupsObj->properties = properties;

	pLvwGroupsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLvwGroupsObj->Release();
	return S_OK;
}

STDMETHODIMP ListViewGroups::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}
	ATLASSUME(properties.pOwnerExLvw);

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		#ifdef USE_STL
			if(properties.nextEnumeratedGroup >= static_cast<int>(properties.pOwnerExLvw->groups.size())) {
		#else
			if(properties.nextEnumeratedGroup >= static_cast<int>(properties.pOwnerExLvw->groups.GetCount())) {
		#endif
			properties.nextEnumeratedGroup = 0;
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitGroup(properties.pOwnerExLvw->groups[properties.nextEnumeratedGroup], properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal), FALSE);
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedGroup;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListViewGroups::Reset(void)
{
	properties.nextEnumeratedGroup = 0;
	return S_OK;
}

STDMETHODIMP ListViewGroups::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedGroup += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


ListViewGroups::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewGroups::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListViewGroups::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewGroups::get_Item(LONG groupIdentifier, GroupIdentifierTypeConstants groupIdentifierType/* = gitID*/, IListViewGroup** ppGroup/* = NULL*/)
{
	ATLASSERT_POINTER(ppGroup, IListViewGroup*);
	if(!ppGroup) {
		return E_POINTER;
	}

	// retrieve the group's ID
	int groupID = -1;
	switch(groupIdentifierType) {
		case gitID:
			groupID = groupIdentifier;
			break;
		case gitIndex:
			if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
				HWND hWndLvw = properties.GetExLvwHWnd();
				ATLASSERT(IsWindow(hWndLvw));

				LVGROUP group = {0};
				group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
				group.mask = LVGF_GROUPID;
				if(SendMessage(hWndLvw, LVM_GETGROUPINFOBYINDEX, groupIdentifier, reinterpret_cast<LPARAM>(&group))) {
					groupID = group.iGroupId;
				}
			}
			break;
		case gitPosition:
			if(properties.pOwnerExLvw) {
				groupID = properties.pOwnerExLvw->PositionIndexToGroupID(groupIdentifier);
			}
			break;
	}

	if(IsValidGroupID(groupID, properties.pOwnerExLvw)) {
		ClassFactory::InitGroup(groupID, properties.pOwnerExLvw, IID_IListViewGroup, reinterpret_cast<LPUNKNOWN*>(ppGroup), FALSE);
		return S_OK;
	}

	if(groupIdentifierType == gitIndex || groupIdentifierType == gitPosition) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP ListViewGroups::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ListViewGroups::Add(BSTR groupHeaderText, LONG groupID, LONG insertAt/* = -1*/, LONG virtualItemCount/* = 1*/, AlignmentConstants headerAlignment/* = alLeft*/, LONG iconIndex/* = -1*/, VARIANT_BOOL collapsible/* = VARIANT_FALSE*/, VARIANT_BOOL collapsed/* = VARIANT_FALSE*/, BSTR groupFooterText/* = L""*/, AlignmentConstants footerAlignment/* = alLeft*/, BSTR subTitleText/* = L""*/, BSTR taskText/* = L""*/, BSTR subsetLinkText/* = L""*/, VARIANT_BOOL subseted/* = VARIANT_FALSE*/, VARIANT_BOOL showHeader/* = VARIANT_TRUE*/, IListViewGroup** ppAddedGroup/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedGroup, IListViewGroup*);
	if(!ppAddedGroup) {
		return E_POINTER;
	}
	ATLASSERT(groupID != -1);
	ATLASSERT(virtualItemCount >= 0);
	if(groupID == -1 || virtualItemCount < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(SendMessage(hWndLvw, LVM_HASGROUP, groupID, 0)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	ATLASSERT(insertAt >= -1);
	if(insertAt < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;

	LVGROUP group = {0};
	group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
	group.iGroupId = groupID;
	group.pszHeader = OLE2W_EX_DEF(groupHeaderText);
	group.cchHeader = lstrlenW(group.pszHeader);
	switch(headerAlignment) {
		case alLeft:
			group.uAlign |= LVGA_HEADER_LEFT;
			break;
		case alCenter:
			group.uAlign |= LVGA_HEADER_CENTER;
			break;
		case alRight:
			group.uAlign |= LVGA_HEADER_RIGHT;
			break;
	}
	group.mask = LVGF_ALIGN | LVGF_GROUPID | LVGF_HEADER;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		group.cItems = virtualItemCount;
		group.mask |= LVGF_ITEMS;
		if(iconIndex != -1) {
			group.iTitleImage = iconIndex;
			group.mask |= LVGF_TITLEIMAGE;
		}
		if(collapsible != VARIANT_FALSE) {
			group.state |= LVGS_COLLAPSIBLE;
			group.stateMask |= LVGS_COLLAPSIBLE;
			group.mask |= LVGF_STATE;
		}
		if(collapsed != VARIANT_FALSE) {
			group.state |= LVGS_COLLAPSED;
			group.stateMask |= LVGS_COLLAPSED;
			group.mask |= LVGF_STATE;
		}
		group.pszFooter = OLE2W_EX_DEF(groupFooterText);
		group.cchFooter = lstrlenW(group.pszFooter);
		if(group.cchFooter > 0) {
			switch(footerAlignment) {
				case alLeft:
					group.uAlign |= LVGA_FOOTER_LEFT;
					break;
				case alCenter:
					group.uAlign |= LVGA_FOOTER_CENTER;
					break;
				case alRight:
					group.uAlign |= LVGA_FOOTER_RIGHT;
					break;
			}
			group.mask |= LVGF_FOOTER;
		} else {
			group.pszFooter = NULL;
		}
		group.pszSubtitle = OLE2W_EX_DEF(subTitleText);
		group.cchSubtitle = lstrlenW(group.pszSubtitle);
		if(group.cchSubtitle > 0) {
			group.mask |= LVGF_SUBTITLE;
		} else {
			group.pszSubtitle = NULL;
		}
		group.pszTask = OLE2W_EX_DEF(taskText);
		group.cchTask = lstrlenW(group.pszTask);
		if(group.cchTask > 0) {
			group.mask |= LVGF_TASK;
		} else {
			group.pszTask = NULL;
		}
		group.pszSubsetTitle = OLE2W_EX_DEF(subsetLinkText);
		group.cchSubsetTitle = lstrlenW(group.pszSubsetTitle);
		if(group.cchSubsetTitle > 0) {
			group.mask |= LVGF_SUBSET;
		} else {
			group.pszSubsetTitle = NULL;
		}
		if(subseted != VARIANT_FALSE) {
			group.state |= LVGS_SUBSETED;
			group.stateMask |= LVGS_SUBSETED;
			group.mask |= LVGF_STATE;
		}
		if(showHeader == VARIANT_FALSE) {
			group.state |= LVGS_NOHEADER;
			group.stateMask |= LVGS_NOHEADER;
			group.mask |= LVGF_STATE;
		}
	}

	// finally insert the group
	*ppAddedGroup = NULL;
	if(SendMessage(hWndLvw, LVM_INSERTGROUP, insertAt, reinterpret_cast<LPARAM>(&group)) != -1) {
		ClassFactory::InitGroup(groupID, properties.pOwnerExLvw, IID_IListViewGroup, reinterpret_cast<LPUNKNOWN*>(ppAddedGroup));
		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP ListViewGroups::Contains(LONG groupIdentifier, GroupIdentifierTypeConstants groupIdentifierType/* = gitID*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the group's ID
	int groupID = -1;
	switch(groupIdentifierType) {
		case gitID:
			groupID = groupIdentifier;
			break;
		case gitIndex:
			if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
				HWND hWndLvw = properties.GetExLvwHWnd();
				ATLASSERT(IsWindow(hWndLvw));

				LVGROUP group = {0};
				group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
				group.mask = LVGF_GROUPID;
				if(SendMessage(hWndLvw, LVM_GETGROUPINFOBYINDEX, groupIdentifier, reinterpret_cast<LPARAM>(&group))) {
					*pValue = VARIANT_TRUE;
				} else {
					*pValue = VARIANT_FALSE;
				}
				return S_OK;
			} else {
				return E_FAIL;
			}
			break;
		case gitPosition:
			if(properties.pOwnerExLvw) {
				groupID = properties.pOwnerExLvw->PositionIndexToGroupID(groupIdentifier);
			}
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsValidGroupID(groupID, properties.pOwnerExLvw));
	return S_OK;
}

STDMETHODIMP ListViewGroups::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		*pValue = static_cast<LONG>(SendMessage(hWndLvw, LVM_GETGROUPCOUNT, 0, 0));
		return S_OK;
	} else {
		// LVM_GETGROUPCOUNT exists in Windows SDK only?!
		#ifdef USE_STL
			*pValue = static_cast<LONG>(properties.pOwnerExLvw->groups.size());
		#else
			*pValue = static_cast<LONG>(properties.pOwnerExLvw->groups.GetCount());
		#endif
		return S_OK;
	}
}

STDMETHODIMP ListViewGroups::Remove(LONG groupIdentifier, GroupIdentifierTypeConstants groupIdentifierType/* = gitID*/)
{
	// get the group's ID
	int groupID = -1;
	switch(groupIdentifierType) {
		case gitID:
			groupID = groupIdentifier;
			break;
		case gitIndex:
			if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
				HWND hWndLvw = properties.GetExLvwHWnd();
				ATLASSERT(IsWindow(hWndLvw));

				LVGROUP group = {0};
				group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
				group.mask = LVGF_GROUPID;
				if(SendMessage(hWndLvw, LVM_GETGROUPINFOBYINDEX, groupIdentifier, reinterpret_cast<LPARAM>(&group))) {
					groupID = group.iGroupId;
				}
			}
			break;
		case gitPosition:
			if(properties.pOwnerExLvw) {
				groupID = properties.pOwnerExLvw->PositionIndexToGroupID(groupIdentifier);
			}
			break;
	}

	if(groupID != -1) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		SendMessage(hWndLvw, LVM_REMOVEGROUP, groupID, 0);
		return S_OK;
	}

	if((groupIdentifierType == gitPosition) || (groupIdentifierType == gitIndex)) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP ListViewGroups::RemoveAll(void)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(SendMessage(hWndLvw, LVM_REMOVEALLGROUPS, 0, 0)) {
		InvalidateRect(hWndLvw, NULL, TRUE);
		return S_OK;
	} else {
		return E_FAIL;
	}
}