// VirtualListViewGroup.cpp: A wrapper for non-existing listview groups.

#include "stdafx.h"
#include "VirtualListViewGroup.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualListViewGroup::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualListViewGroup, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualListViewGroup::Properties::~Properties()
{
	if(settings.pszDescriptionBottom) {
		SECUREFREE(settings.pszDescriptionBottom);
	}
	if(settings.pszDescriptionTop) {
		SECUREFREE(settings.pszDescriptionTop);
	}
	if(settings.pszFooter) {
		SECUREFREE(settings.pszFooter);
	}
	if(settings.pszHeader) {
		SECUREFREE(settings.pszHeader);
	}
	if(settings.pszSubsetTitle) {
		SECUREFREE(settings.pszSubsetTitle);
	}
	if(settings.pszSubtitle) {
		SECUREFREE(settings.pszSubtitle);
	}
	if(settings.pszTask) {
		SECUREFREE(settings.pszTask);
	}
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}


HRESULT VirtualListViewGroup::LoadState(PLVGROUP pSource)
{
	ATLASSERT_POINTER(pSource, LVGROUP);

	SECUREFREE(properties.settings.pszDescriptionBottom);
	SECUREFREE(properties.settings.pszDescriptionTop);
	SECUREFREE(properties.settings.pszFooter);
	SECUREFREE(properties.settings.pszHeader);
	SECUREFREE(properties.settings.pszSubsetTitle);
	SECUREFREE(properties.settings.pszSubtitle);
	SECUREFREE(properties.settings.pszTask);
	properties.settings = *pSource;
	if(properties.settings.mask & LVGF_DESCRIPTIONBOTTOM) {
		// duplicate the description text
		properties.settings.cchDescriptionBottom = lstrlenW(pSource->pszDescriptionBottom);
		properties.settings.pszDescriptionBottom = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchDescriptionBottom + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszDescriptionBottom, properties.settings.cchDescriptionBottom + 1, pSource->pszDescriptionBottom)));
	}
	if(properties.settings.mask & LVGF_DESCRIPTIONTOP) {
		// duplicate the description text
		properties.settings.cchDescriptionTop = lstrlenW(pSource->pszDescriptionTop);
		properties.settings.pszDescriptionTop = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchDescriptionTop + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszDescriptionTop, properties.settings.cchDescriptionTop + 1, pSource->pszDescriptionTop)));
	}
	if(properties.settings.mask & LVGF_FOOTER) {
		// duplicate the footer text
		properties.settings.cchFooter = lstrlenW(pSource->pszFooter);
		properties.settings.pszFooter = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchFooter + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszFooter, properties.settings.cchFooter + 1, pSource->pszFooter)));
	}
	if(properties.settings.mask & LVGF_HEADER) {
		// duplicate the header text
		properties.settings.cchHeader = lstrlenW(pSource->pszHeader);
		properties.settings.pszHeader = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchHeader + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszHeader, properties.settings.cchHeader + 1, pSource->pszHeader)));
	}
	if(properties.settings.mask & LVGF_SUBSET) {
		// duplicate the subset link text
		properties.settings.cchSubsetTitle = lstrlenW(pSource->pszSubsetTitle);
		properties.settings.pszSubsetTitle = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchSubsetTitle + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszSubsetTitle, properties.settings.cchSubsetTitle + 1, pSource->pszSubsetTitle)));
	}
	if(properties.settings.mask & LVGF_SUBTITLE) {
		// duplicate the subtitle text
		properties.settings.cchSubtitle = lstrlenW(pSource->pszSubtitle);
		properties.settings.pszSubtitle = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchSubtitle + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszSubtitle, properties.settings.cchSubtitle + 1, pSource->pszSubtitle)));
	}
	if(properties.settings.mask & LVGF_TASK) {
		// duplicate the task text
		properties.settings.cchTask = lstrlenW(pSource->pszTask);
		properties.settings.pszTask = reinterpret_cast<LPWSTR>(malloc((properties.settings.cchTask + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.settings.pszTask, properties.settings.cchTask + 1, pSource->pszTask)));
	}

	return S_OK;
}

HRESULT VirtualListViewGroup::SaveState(PLVGROUP pTarget)
{
	ATLASSERT_POINTER(pTarget, LVGROUP);

	SECUREFREE(pTarget->pszDescriptionBottom);
	SECUREFREE(pTarget->pszDescriptionTop);
	SECUREFREE(pTarget->pszFooter);
	SECUREFREE(pTarget->pszHeader);
	SECUREFREE(pTarget->pszSubsetTitle);
	SECUREFREE(pTarget->pszSubtitle);
	SECUREFREE(pTarget->pszTask);
	*pTarget = properties.settings;
	if(pTarget->mask & LVGF_DESCRIPTIONBOTTOM) {
		// duplicate the description text
		pTarget->pszDescriptionBottom = reinterpret_cast<LPWSTR>(malloc((pTarget->cchDescriptionBottom + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszDescriptionBottom, pTarget->cchDescriptionBottom + 1, properties.settings.pszDescriptionBottom)));
	}
	if(pTarget->mask & LVGF_DESCRIPTIONTOP) {
		// duplicate the description text
		pTarget->pszDescriptionTop = reinterpret_cast<LPWSTR>(malloc((pTarget->cchDescriptionTop + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszDescriptionTop, pTarget->cchDescriptionTop + 1, properties.settings.pszDescriptionTop)));
	}
	if(pTarget->mask & LVGF_FOOTER) {
		// duplicate the footer text
		pTarget->pszFooter = reinterpret_cast<LPWSTR>(malloc((pTarget->cchFooter + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszFooter, pTarget->cchFooter + 1, properties.settings.pszFooter)));
	}
	if(pTarget->mask & LVGF_HEADER) {
		// duplicate the header text
		pTarget->pszHeader = reinterpret_cast<LPWSTR>(malloc((pTarget->cchHeader + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszHeader, pTarget->cchHeader + 1, properties.settings.pszHeader)));
	}
	if(pTarget->mask & LVGF_SUBSET) {
		// duplicate the subset link text
		pTarget->pszSubsetTitle = reinterpret_cast<LPWSTR>(malloc((pTarget->cchSubsetTitle + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszSubsetTitle, pTarget->cchSubsetTitle + 1, properties.settings.pszSubsetTitle)));
	}
	if(pTarget->mask & LVGF_SUBTITLE) {
		// duplicate the subtitle text
		pTarget->pszSubtitle = reinterpret_cast<LPWSTR>(malloc((pTarget->cchSubtitle + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszSubtitle, pTarget->cchSubtitle + 1, properties.settings.pszSubtitle)));
	}
	if(pTarget->mask & LVGF_TASK) {
		// duplicate the task text
		pTarget->pszTask = reinterpret_cast<LPWSTR>(malloc((pTarget->cchTask + 1) * sizeof(WCHAR)));
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pTarget->pszTask, pTarget->cchTask + 1, properties.settings.pszTask)));
	}

	return S_OK;
}

void VirtualListViewGroup::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP VirtualListViewGroup::get_Alignment(GroupComponentConstants groupComponent/* = gcHeader*/, AlignmentConstants* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_ALIGN) {
		switch(groupComponent) {
			case gcHeader:
				switch(properties.settings.uAlign & (LVGA_HEADER_LEFT | LVGA_HEADER_CENTER | LVGA_HEADER_RIGHT)) {
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
				switch(properties.settings.uAlign & (LVGA_FOOTER_LEFT | LVGA_FOOTER_CENTER | LVGA_FOOTER_RIGHT)) {
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
	} else {
		*pValue = alLeft;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVGS_FOCUSED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_Collapsed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVGS_COLLAPSED);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_Collapsible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVGS_COLLAPSIBLE);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_DescriptionTextBottom(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_DESCRIPTIONBOTTOM) {
		*pValue = _bstr_t(properties.settings.pszDescriptionBottom).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_DescriptionTextTop(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_DESCRIPTIONTOP) {
		*pValue = _bstr_t(properties.settings.pszDescriptionTop).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_TITLEIMAGE) {
		*pValue = properties.settings.iTitleImage;
	} else {
		*pValue = -1;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = -1;
	if(properties.settings.mask & LVGF_GROUPID) {
		*pValue = properties.settings.iGroupId;
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVGS_SELECTED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_ShowHeader(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL((properties.settings.state & LVGS_NOHEADER) == 0);
	} else {
		*pValue = VARIANT_TRUE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_Subseted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVGS_SUBSETED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_SubsetLinkFocused(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVGS_SUBSETLINKFOCUSED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_SubsetLinkText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_SUBSET) {
		*pValue = _bstr_t(properties.settings.pszSubsetTitle).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_SubtitleText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_SUBTITLE) {
		*pValue = _bstr_t(properties.settings.pszSubtitle).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_TaskText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVGF_TASK) {
		*pValue = _bstr_t(properties.settings.pszTask).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewGroup::get_Text(GroupComponentConstants groupComponent/* = gcHeader*/, BSTR* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"");
	switch(groupComponent) {
		case gcHeader:
			if(properties.settings.mask & LVGF_HEADER) {
				*pValue = _bstr_t(properties.settings.pszHeader).Detach();
			}
			break;
		case gcFooter:
			if(properties.settings.mask & LVGF_FOOTER) {
				*pValue = _bstr_t(properties.settings.pszFooter).Detach();
			}
			break;
	}

	return S_OK;
}