// ListViewGroup.cpp: A wrapper for existing listview groups.

#include "stdafx.h"
#include "ListViewGroup.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewGroup::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewGroup, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListViewGroup::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewGroup::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT ListViewGroup::SaveState(int groupID, PLVGROUP pTarget, HWND hWndLvw/* = NULL*/)
{
	if(!hWndLvw) {
		hWndLvw = properties.GetExLvwHWnd();
	}
	ATLASSERT(IsWindow(hWndLvw));
	if(!IsValidGroupID(groupID, hWndLvw)) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, LVGROUP);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(LVGROUP));
	pTarget->iGroupId = groupID;
	pTarget->mask = LVGF_ALIGN | LVGF_FOOTER | LVGF_GROUPID | LVGF_HEADER | LVGF_STATE;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		pTarget->mask |= LVGF_DESCRIPTIONBOTTOM | LVGF_DESCRIPTIONTOP | LVGF_SUBTITLE | LVGF_TASK | LVGF_TITLEIMAGE;
		pTarget->cchDescriptionBottom = MAX_GROUPDESCRIPTIONTEXTLENGTH;
		pTarget->pszDescriptionBottom = reinterpret_cast<LPWSTR>(malloc((pTarget->cchDescriptionBottom + 1) * sizeof(WCHAR)));
		pTarget->cchDescriptionTop = MAX_GROUPDESCRIPTIONTEXTLENGTH;
		pTarget->pszDescriptionTop = reinterpret_cast<LPWSTR>(malloc((pTarget->cchDescriptionTop + 1) * sizeof(WCHAR)));
		pTarget->cchSubtitle = MAX_GROUPSUBTITLETEXTLENGTH;
		pTarget->pszSubtitle = reinterpret_cast<LPWSTR>(malloc((pTarget->cchSubtitle + 1) * sizeof(WCHAR)));
		pTarget->cchTask = MAX_GROUPTASKTEXTLENGTH;
		pTarget->pszTask = reinterpret_cast<LPWSTR>(malloc((pTarget->cchTask + 1) * sizeof(WCHAR)));
	}

	return static_cast<HRESULT>(SendMessage(hWndLvw, LVM_GETGROUPINFO, groupID, reinterpret_cast<LPARAM>(pTarget)));
}


void ListViewGroup::Attach(int groupID)
{
	properties.groupID = groupID;
}

void ListViewGroup::Detach(void)
{
	properties.groupID = -1;
}

HRESULT ListViewGroup::LoadState(PLVGROUP pSource)
{
	ATLASSERT_POINTER(pSource, LVGROUP);
	if(!pSource) {
		return E_POINTER;
	}

	// now apply the settings
	AlignmentConstants alignment = alLeft;
	if(pSource->mask & LVGF_ALIGN) {
		switch(pSource->uAlign & (LVGA_HEADER_LEFT | LVGA_HEADER_CENTER | LVGA_HEADER_RIGHT)) {
			case LVGA_HEADER_LEFT:
				alignment = alLeft;
				break;
			case LVGA_HEADER_CENTER:
				alignment = alCenter;
				break;
			case LVGA_HEADER_RIGHT:
				alignment = alRight;
				break;
		}
		put_Alignment(gcHeader, alignment);
		alignment = alLeft;
		switch(pSource->uAlign & (LVGA_FOOTER_LEFT | LVGA_FOOTER_CENTER | LVGA_FOOTER_RIGHT)) {
			case LVGA_FOOTER_LEFT:
				alignment = alLeft;
				break;
			case LVGA_FOOTER_CENTER:
				alignment = alCenter;
				break;
			case LVGA_FOOTER_RIGHT:
				alignment = alRight;
				break;
		}
		put_Alignment(gcFooter, alignment);
	}

	if(pSource->mask & LVGF_HEADER) {
		BSTR t = SysAllocString(L"");
		put_Text(gcHeader, _bstr_t(pSource->pszHeader).Detach());
		SysFreeString(t);
	}
	if(pSource->mask & LVGF_FOOTER) {
		BSTR t = SysAllocString(L"");
		put_Text(gcFooter, _bstr_t(pSource->pszFooter).Detach());
		SysFreeString(t);
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		VARIANT_BOOL b = VARIANT_FALSE;
		if(pSource->mask & LVGF_STATE) {
			b = BOOL2VARIANTBOOL(pSource->state & LVGS_COLLAPSED);
			put_Collapsed(b);
			b = BOOL2VARIANTBOOL(pSource->state & LVGS_COLLAPSIBLE);
			put_Collapsible(b);
			b = BOOL2VARIANTBOOL(pSource->state & LVGS_SELECTED);
			put_Selected(b);
			b = BOOL2VARIANTBOOL((pSource->state & LVGS_NOHEADER) == 0);
			put_ShowHeader(b);
			b = BOOL2VARIANTBOOL(pSource->state & LVGS_SUBSETED);
			put_Subseted(b);
			b = BOOL2VARIANTBOOL(pSource->state & LVGS_SUBSETLINKFOCUSED);
			put_SubsetLinkFocused(b);
		}

		LONG l = -1;
		if(pSource->mask & LVGF_TITLEIMAGE) {
			l = pSource->iTitleImage;
			put_IconIndex(l);
		}

		if(pSource->mask & LVGF_DESCRIPTIONBOTTOM) {
			BSTR t = SysAllocString(L"");
			put_DescriptionTextBottom(_bstr_t(pSource->pszDescriptionBottom).Detach());
			SysFreeString(t);
		}
		if(pSource->mask & LVGF_DESCRIPTIONTOP) {
			BSTR t = SysAllocString(L"");
			put_DescriptionTextTop(_bstr_t(pSource->pszDescriptionTop).Detach());
			SysFreeString(t);
		}
		if(pSource->mask & LVGF_SUBTITLE) {
			BSTR t = SysAllocString(L"");
			put_SubtitleText(_bstr_t(pSource->pszSubtitle).Detach());
			SysFreeString(t);
		}
		if(pSource->mask & LVGF_TASK) {
			BSTR t = SysAllocString(L"");
			put_TaskText(_bstr_t(pSource->pszTask).Detach());
			SysFreeString(t);
		}
	}

	return S_OK;
}

HRESULT ListViewGroup::LoadState(VirtualListViewGroup* pSource)
{
	ATLASSUME(pSource);
	if(!pSource) {
		return E_POINTER;
	}

	// now apply the settings
	AlignmentConstants alignment = alLeft;
	pSource->get_Alignment(gcHeader, &alignment);
	put_Alignment(gcHeader, alignment);
	alignment = alLeft;
	pSource->get_Alignment(gcFooter, &alignment);
	put_Alignment(gcFooter, alignment);

	BSTR t = SysAllocString(L"");
	pSource->get_Text(gcHeader, &t);
	put_Text(gcHeader, t);
	SysFreeString(t);
	t = SysAllocString(L"");
	pSource->get_Text(gcFooter, &t);
	put_Text(gcFooter, t);
	SysFreeString(t);

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		VARIANT_BOOL b = VARIANT_FALSE;
		pSource->get_Collapsed(&b);
		put_Collapsed(b);
		b = VARIANT_FALSE;
		pSource->get_Collapsible(&b);
		put_Collapsible(b);
		b = VARIANT_FALSE;
		pSource->get_Selected(&b);
		put_Selected(b);
		b = VARIANT_TRUE;
		pSource->get_ShowHeader(&b);
		put_ShowHeader(b);
		b = VARIANT_FALSE;
		pSource->get_Subseted(&b);
		put_Subseted(b);
		b = VARIANT_FALSE;
		pSource->get_SubsetLinkFocused(&b);
		put_SubsetLinkFocused(b);

		LONG l = -1;
		pSource->get_IconIndex(&l);
		put_IconIndex(l);

		t = SysAllocString(L"");
		pSource->get_DescriptionTextBottom(&t);
		put_DescriptionTextBottom(t);
		SysFreeString(t);
		t = SysAllocString(L"");
		pSource->get_DescriptionTextTop(&t);
		put_DescriptionTextTop(t);
		SysFreeString(t);
		t = SysAllocString(L"");
		pSource->get_SubtitleText(&t);
		put_SubtitleText(t);
		SysFreeString(t);
		t = SysAllocString(L"");
		pSource->get_TaskText(&t);
		put_TaskText(t);
		SysFreeString(t);
	}

	return S_OK;
}

HRESULT ListViewGroup::SaveState(PLVGROUP pTarget, HWND hWndLvw/* = NULL*/)
{
	return SaveState(properties.groupID, pTarget, hWndLvw);
}

HRESULT ListViewGroup::SaveState(VirtualListViewGroup* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerExLvw);
	LVGROUP group = {0};
	HRESULT hr = SaveState(&group);
	pTarget->LoadState(&group);
	SECUREFREE(group.pszDescriptionBottom);
	SECUREFREE(group.pszDescriptionTop);
	SECUREFREE(group.pszFooter);
	SECUREFREE(group.pszHeader);
	SECUREFREE(group.pszSubtitle);
	SECUREFREE(group.pszTask);

	return hr;
}

void ListViewGroup::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewGroup::get_Alignment(GroupComponentConstants groupComponent/* = gcHeader*/, AlignmentConstants* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVGROUP group = {0};
	group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
	group.mask = LVGF_ALIGN;

	if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
		switch(groupComponent) {
			case gcHeader:
				switch(group.uAlign & (LVGA_HEADER_LEFT | LVGA_HEADER_CENTER | LVGA_HEADER_RIGHT)) {
					case LVGA_HEADER_LEFT:
						*pValue = alLeft;
						break;
					case LVGA_HEADER_CENTER:
						*pValue = alCenter;
						break;
					case LVGA_HEADER_RIGHT:
						*pValue = alRight;
						break;
				}
				break;
			case gcFooter:
				switch(group.uAlign & (LVGA_FOOTER_LEFT | LVGA_FOOTER_CENTER | LVGA_FOOTER_RIGHT)) {
					case LVGA_FOOTER_LEFT:
						*pValue = alLeft;
						break;
					case LVGA_FOOTER_CENTER:
						*pValue = alCenter;
						break;
					case LVGA_FOOTER_RIGHT:
						*pValue = alRight;
						break;
				}
				break;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_Alignment(GroupComponentConstants groupComponent/* = gcHeader*/, AlignmentConstants newValue/* = alLeft*/)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVGROUP group = {0};
	group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
	// the next line is required on Windows XP - don't remove it
	group.iGroupId = properties.groupID;
	group.mask = LVGF_ALIGN;

	switch(groupComponent) {
		case gcHeader:
			switch(newValue) {
				case alLeft:
					group.uAlign = LVGA_HEADER_LEFT;
					break;
				case alCenter:
					group.uAlign = LVGA_HEADER_CENTER;
					break;
				case alRight:
					group.uAlign = LVGA_HEADER_RIGHT;
					break;
			}
			break;
		case gcFooter:
			switch(newValue) {
				case alLeft:
					group.uAlign = LVGA_FOOTER_LEFT;
					break;
				case alCenter:
					group.uAlign = LVGA_FOOTER_CENTER;
					break;
				case alRight:
					group.uAlign = LVGA_FOOTER_RIGHT;
					break;
			}
			break;
	}

	if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
		InvalidateRect(hWndLvw, NULL, TRUE);
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_FOCUSED));
		*pValue = BOOL2VARIANTBOOL(state & LVGS_FOCUSED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Collapsed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_COLLAPSED));
		*pValue = BOOL2VARIANTBOOL(state & LVGS_COLLAPSED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_Collapsed(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newValue != VARIANT_FALSE) {
			group.state = LVGS_COLLAPSED;
		}
		group.stateMask = LVGS_COLLAPSED;
		group.mask = LVGF_STATE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Collapsible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_COLLAPSIBLE));
		*pValue = BOOL2VARIANTBOOL(state & LVGS_COLLAPSIBLE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_Collapsible(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newValue != VARIANT_FALSE) {
			group.state = LVGS_COLLAPSIBLE;
		}
		group.stateMask = LVGS_COLLAPSIBLE;
		group.mask = LVGF_STATE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_DescriptionTextBottom(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.cchDescriptionBottom = MAX_GROUPDESCRIPTIONTEXTLENGTH;
		group.pszDescriptionBottom = reinterpret_cast<LPWSTR>(malloc((group.cchDescriptionBottom + 1) * sizeof(WCHAR)));
		group.mask = LVGF_DESCRIPTIONBOTTOM;

		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = _bstr_t(group.pszDescriptionBottom).Detach();
			hr = S_OK;
		}
		SECUREFREE(group.pszDescriptionBottom);
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_DescriptionTextBottom(BSTR newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.pszDescriptionBottom = OLE2W_EX_DEF(newValue);
		group.cchDescriptionBottom = lstrlenW(group.pszDescriptionBottom);
		group.mask = LVGF_DESCRIPTIONBOTTOM;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_DescriptionTextTop(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.cchDescriptionTop = MAX_GROUPDESCRIPTIONTEXTLENGTH;
		group.pszDescriptionTop = reinterpret_cast<LPWSTR>(malloc((group.cchDescriptionTop + 1) * sizeof(WCHAR)));
		group.mask = LVGF_DESCRIPTIONTOP;

		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = _bstr_t(group.pszDescriptionTop).Detach();
			hr = S_OK;
		}
		SECUREFREE(group.pszDescriptionTop);
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_DescriptionTextTop(BSTR newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.pszDescriptionTop = OLE2W_EX_DEF(newValue);
		group.cchDescriptionTop = lstrlenW(group.pszDescriptionTop);
		group.mask = LVGF_DESCRIPTIONTOP;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_FirstItem(IListViewItem** ppItem)
{
	ATLASSERT_POINTER(ppItem, IListViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.mask = LVGF_ITEMS;

		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			LVITEMINDEX itemIndex = {group.iFirstItem, 0};
			ClassFactory::InitListItem(itemIndex, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewGroup::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.mask = LVGF_TITLEIMAGE;

		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = group.iTitleImage;
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_IconIndex(LONG newValue)
{
	if(newValue < -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.iTitleImage = newValue;
		group.mask = LVGF_TITLEIMAGE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.groupID;
	return S_OK;
}

STDMETHODIMP ListViewGroup::put_ID(LONG newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVGROUP group = {0};
	group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		group.iGroupId = newValue;
		group.mask = LVGF_GROUPID;
	} else {
		/* LVM_GETGROUPINFO is a bit broken under Windows XP. Changing the group's ID changes the group's
		   alignment. */
		group.iGroupId = properties.groupID;
		group.mask = LVGF_ALIGN | LVGF_GROUPID | LVGF_STATE;
		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) != properties.groupID) {
			return E_FAIL;
		}
		group.iGroupId = newValue;
	}

	if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
		/* TODO: Remove the next line if we ever keep a list of all ListViewGroup objects to update their
		         IDs on LVM_SETGROUPINFO. */
		properties.groupID = newValue;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.mask = LVGF_GROUPID;
		int groups = static_cast<int>(SendMessage(hWndLvw, LVM_GETGROUPCOUNT, 0, 0));
		for(int groupIndex = 0; groupIndex < groups; ++groupIndex) {
			if(SendMessage(hWndLvw, LVM_GETGROUPINFOBYINDEX, groupIndex, reinterpret_cast<LPARAM>(&group))) {
				if(group.iGroupId == properties.groupID) {
					*pValue = groupIndex;
					return S_OK;
				}
			}
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_ItemCount(VARIANT_BOOL visibleSubsetOnly/* = VARIANT_FALSE*/, LONG* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.mask = (visibleSubsetOnly == VARIANT_FALSE ? LVGF_ITEMS : LVGF_SUBSETITEMS);

		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = group.cItems;
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_ItemCount(VARIANT_BOOL visibleSubsetOnly/* = VARIANT_FALSE*/, LONG newValue/* = 0*/)
{
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.cItems = newValue;
		group.mask = (visibleSubsetOnly == VARIANT_FALSE ? LVGF_ITEMS : LVGF_SUBSETITEMS);

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_ListItems(IListViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IListViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitListItems(properties.pOwnerExLvw, IID_IListViewItems, reinterpret_cast<LPUNKNOWN*>(ppItems));

	if(*ppItems) {
		// install a fpGroup filter
		VARIANT filter;
		VariantInit(&filter);
		// create the array
		filter.vt = VT_ARRAY | VT_VARIANT;
		filter.parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		// store properties.groupID in the array (as object)
		VARIANT element;
		VariantInit(&element);
		element.vt = VT_DISPATCH;
		ClassFactory::InitGroup(properties.groupID, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&element.pdispVal));
		LONG index = 1;
		SafeArrayPutElement(filter.parray, &index, &element);
		VariantClear(&element);

		(*ppItems)->put_FilterType(fpGroup, ftIncluding);
		(*ppItems)->put_Filter(fpGroup, filter);
		VariantClear(&filter);
	}

	return S_OK;
}

STDMETHODIMP ListViewGroup::get_Position(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw) {
		if((*pValue = properties.pOwnerExLvw->GroupIDToPositionIndex(properties.groupID)) != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_SELECTED));
		*pValue = BOOL2VARIANTBOOL(state & LVGS_SELECTED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_Selected(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newValue != VARIANT_FALSE) {
			group.state = LVGS_SELECTED;
		}
		group.stateMask = LVGS_SELECTED;
		group.mask = LVGF_STATE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_ShowHeader(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_NOHEADER));
		*pValue = BOOL2VARIANTBOOL((state & LVGS_NOHEADER) == 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_ShowHeader(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newValue == VARIANT_FALSE) {
			group.state = LVGS_NOHEADER;
		}
		group.stateMask = LVGS_NOHEADER;
		group.mask = LVGF_STATE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Subseted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_SUBSETED));
		*pValue = BOOL2VARIANTBOOL(state & LVGS_SUBSETED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_Subseted(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newValue != VARIANT_FALSE) {
			group.state = LVGS_SUBSETED;
		}
		group.stateMask = LVGS_SUBSETED;
		group.mask = LVGF_STATE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_SubsetLinkFocused(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		UINT state = static_cast<UINT>(SendMessage(hWndLvw, LVM_GETGROUPSTATE, properties.groupID, LVGS_SUBSETLINKFOCUSED));
		*pValue = BOOL2VARIANTBOOL(state & LVGS_SUBSETLINKFOCUSED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_SubsetLinkFocused(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newValue != VARIANT_FALSE) {
			group.state = LVGS_SUBSETLINKFOCUSED;
		}
		group.stateMask = LVGS_SUBSETLINKFOCUSED;
		group.mask = LVGF_STATE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_SubsetLinkText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.cchSubsetTitle = MAX_GROUPSUBSETLINKTEXTLENGTH;
		group.pszSubsetTitle = reinterpret_cast<LPWSTR>(malloc((group.cchSubsetTitle + 1) * sizeof(WCHAR)));
		group.mask = LVGF_SUBSET;

		BOOL caughtAnAccessViolation = FALSE;
		if(GuardedSendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group), &caughtAnAccessViolation) == properties.groupID) {
		//if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = _bstr_t(group.pszSubsetTitle).Detach();
			hr = S_OK;
		} else if(caughtAnAccessViolation) {
			*pValue = _bstr_t(TEXT("")).Detach();
			hr = S_OK;
		}
		SECUREFREE(group.pszSubsetTitle);
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_SubsetLinkText(BSTR newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.pszSubsetTitle = OLE2W_EX_DEF(newValue);
		group.cchSubsetTitle = lstrlenW(group.pszSubsetTitle);
		group.mask = LVGF_SUBSET;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_SubtitleText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.cchSubtitle = MAX_GROUPSUBTITLETEXTLENGTH;
		group.pszSubtitle = reinterpret_cast<LPWSTR>(malloc((group.cchSubtitle + 1) * sizeof(WCHAR)));
		group.mask = LVGF_SUBTITLE;

		BOOL caughtAnAccessViolation = FALSE;
		if(GuardedSendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group), &caughtAnAccessViolation) == properties.groupID) {
		//if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = _bstr_t(group.pszSubtitle).Detach();
			hr = S_OK;
		} else if(caughtAnAccessViolation) {
			*pValue = _bstr_t(TEXT("")).Detach();
			hr = S_OK;
		}
		SECUREFREE(group.pszSubtitle);
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_SubtitleText(BSTR newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.pszSubtitle = OLE2W_EX_DEF(newValue);
		group.cchSubtitle = lstrlenW(group.pszSubtitle);
		group.mask = LVGF_SUBTITLE;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_TaskText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.cchTask = MAX_GROUPTASKTEXTLENGTH;
		group.pszTask = reinterpret_cast<LPWSTR>(malloc((group.cchTask + 1) * sizeof(WCHAR)));
		group.mask = LVGF_TASK;

		BOOL caughtAnAccessViolation = FALSE;
		if(GuardedSendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group), &caughtAnAccessViolation) == properties.groupID) {
		//if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			*pValue = _bstr_t(group.pszTask).Detach();
			hr = S_OK;
		} else if(caughtAnAccessViolation) {
			*pValue = _bstr_t(TEXT("")).Detach();
			hr = S_OK;
		}
		SECUREFREE(group.pszTask);
	}
	return hr;
}

STDMETHODIMP ListViewGroup::put_TaskText(BSTR newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.pszTask = OLE2W_EX_DEF(newValue);
		group.cchTask = lstrlenW(group.pszTask);
		group.mask = LVGF_TASK;

		if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
			InvalidateRect(hWndLvw, NULL, TRUE);
			return S_OK;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewGroup::get_Text(GroupComponentConstants groupComponent/* = gcHeader*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVGROUP group = {0};
	group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
	// the listview will allocate a string buffer itself
	switch(groupComponent) {
		case gcHeader:
			group.mask = LVGF_HEADER;
			break;
		case gcFooter:
			group.mask = LVGF_FOOTER;
			break;
	}

	if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
		switch(groupComponent) {
			case gcHeader:
				*pValue = _bstr_t(group.pszHeader).Detach();
				break;
			case gcFooter:
				*pValue = _bstr_t(group.pszFooter).Detach();
				break;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewGroup::put_Text(GroupComponentConstants groupComponent/* = gcHeader*/, BSTR newValue/* = L""*/)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVGROUP group = {0};
	group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
	if(!properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		/* Under Windows XP you can't change the group text without changing the group's ID.
		   Also LVM_GETGROUPINFO is a bit broken under Windows XP. Changing the alignment won't change the
		   text, but changing the text changes the alignment. */
		group.iGroupId = properties.groupID;
		group.mask = LVGF_ALIGN | LVGF_GROUPID | LVGF_STATE;
		if(SendMessage(hWndLvw, LVM_GETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) != properties.groupID) {
			return E_FAIL;
		}
	}

	CW2W converter(newValue);
	switch(groupComponent) {
		case gcHeader:
			group.pszHeader = converter;
			group.cchHeader = lstrlenW(group.pszHeader);
			group.mask |= LVGF_HEADER;
			break;
		case gcFooter:
			group.pszFooter = converter;
			group.cchFooter = lstrlenW(group.pszFooter);
			group.mask |= LVGF_FOOTER;
			break;
	}

	if(SendMessage(hWndLvw, LVM_SETGROUPINFO, properties.groupID, reinterpret_cast<LPARAM>(&group)) == properties.groupID) {
		InvalidateRect(hWndLvw, NULL, TRUE);
		return S_OK;
	}

	return E_FAIL;
}


STDMETHODIMP ListViewGroup::GetRectangle(GroupRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		RECT groupRectangle = {0};
		groupRectangle.top = rectangleType;
		if(SendMessage(hWndLvw, LVM_GETGROUPRECT, properties.groupID, reinterpret_cast<LPARAM>(&groupRectangle))) {
			if(pXLeft) {
				*pXLeft = groupRectangle.left;
			}
			if(pYTop) {
				*pYTop = groupRectangle.top;
			}
			if(pXRight) {
				*pXRight = groupRectangle.right;
			}
			if(pYBottom) {
				*pYBottom = groupRectangle.bottom;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}