// VirtualListViewItem.cpp: A wrapper for non-existing listview items.

#include "stdafx.h"
#include "VirtualListViewItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualListViewItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualListViewItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualListViewItem::Properties::~Properties()
{
	if(settings.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(settings.pszText);
	}
	if(settings.puColumns) {
		delete[] settings.puColumns;
		settings.puColumns = NULL;
	}
	if(settings.piColFmt) {
		delete[] settings.piColFmt;
		settings.piColFmt = NULL;
	}
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}


HRESULT VirtualListViewItem::LoadState(LPLVITEM pSource)
{
	ATLASSERT_POINTER(pSource, LVITEM);

	SECUREFREE(properties.settings.pszText);
	if(properties.settings.puColumns) {
		delete[] properties.settings.puColumns;
		properties.settings.puColumns = NULL;
	}
	if(properties.settings.piColFmt) {
		delete[] properties.settings.piColFmt;
		properties.settings.piColFmt = NULL;
	}
	properties.settings = *pSource;
	if(properties.settings.mask & LVIF_TEXT) {
		// duplicate the item's text
		if(properties.settings.pszText != LPSTR_TEXTCALLBACK) {
			properties.settings.cchTextMax = lstrlen(pSource->pszText);
			properties.settings.pszText = reinterpret_cast<LPTSTR>(malloc((properties.settings.cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(properties.settings.pszText, properties.settings.cchTextMax + 1, pSource->pszText)));
		}
	}
	if(properties.settings.mask & LVIF_COLUMNS) {
		// duplicate the array
		if((properties.settings.cColumns > 0) && (properties.settings.cColumns != I_COLUMNSCALLBACK)) {
			properties.settings.puColumns = new UINT[properties.settings.cColumns];
			CopyMemory(properties.settings.puColumns, pSource->puColumns, properties.settings.cColumns * sizeof(UINT));
		}
	}
	// NOTE: If we enable the check for comctl32 6.10, pSource->piColFmt will be freed on comctl32 6.0!
	//if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		if(properties.settings.mask & LVIF_COLFMT) {
			// duplicate the array
			if((properties.settings.cColumns > 0) && (properties.settings.cColumns != I_COLUMNSCALLBACK)) {
				properties.settings.piColFmt = new int[properties.settings.cColumns];
				CopyMemory(properties.settings.piColFmt, pSource->piColFmt, properties.settings.cColumns * sizeof(int));
			}
		}
	//}

	return S_OK;
}

HRESULT VirtualListViewItem::SaveState(LPLVITEM pTarget)
{
	ATLASSERT_POINTER(pTarget, LVITEM);

	SECUREFREE(pTarget->pszText);
	if(pTarget->puColumns) {
		delete[] pTarget->puColumns;
		pTarget->puColumns = NULL;
	}
	if(pTarget->piColFmt) {
		delete[] pTarget->piColFmt;
		pTarget->piColFmt = NULL;
	}
	*pTarget = properties.settings;
	if(pTarget->mask & LVIF_TEXT) {
		// duplicate the item's text
		if(pTarget->pszText != LPSTR_TEXTCALLBACK) {
			pTarget->pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(pTarget->pszText, pTarget->cchTextMax + 1, properties.settings.pszText)));
		}
	}
	if(pTarget->mask & LVIF_COLUMNS) {
		// duplicate the array
		if((pTarget->cColumns > 0) && (pTarget->cColumns != I_COLUMNSCALLBACK)) {
			pTarget->puColumns = new UINT[pTarget->cColumns];
			CopyMemory(pTarget->puColumns, properties.settings.puColumns, pTarget->cColumns * sizeof(UINT));
		}
	}
	// NOTE: If we enable the check for comctl32 6.10, properties.settings.piColFmt will be freed on comctl32 6.0!
	//if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		if(pTarget->mask & LVIF_COLFMT) {
			// duplicate the array
			if((pTarget->cColumns > 0) && (pTarget->cColumns != I_COLUMNSCALLBACK)) {
				pTarget->piColFmt = new int[pTarget->cColumns];
				CopyMemory(pTarget->piColFmt, properties.settings.piColFmt, pTarget->cColumns * sizeof(int));
			}
		}
	//}

	return S_OK;
}

void VirtualListViewItem::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP VirtualListViewItem::get_Activating(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVIS_ACTIVATING);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVIS_FOCUSED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_DropHilited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVIS_DROPHILITED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Ghosted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVIS_CUT);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Glowing(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVIS_GLOW);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Group(IListViewGroup** ppGroup)
{
	ATLASSERT_POINTER(ppGroup, IListViewGroup*);
	if(!ppGroup) {
		return E_POINTER;
	}

	*ppGroup = NULL;
	if(properties.settings.mask & LVIF_GROUPID) {
		ClassFactory::InitGroup(properties.settings.iGroupId, properties.pOwnerExLvw, IID_IListViewGroup, reinterpret_cast<LPUNKNOWN*>(ppGroup));
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_IMAGE) {
		*pValue = properties.settings.iImage;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Indent(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_INDENT) {
		*pValue = properties.settings.iIndent;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.settings.iItem;
	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_PARAM) {
		*pValue = static_cast<LONG>(properties.settings.lParam);
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = ((properties.settings.state & LVIS_OVERLAYMASK) >> 8);
	} else {
		*pValue = 1;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.state & LVIS_SELECTED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_STATE) {
		*pValue = ((properties.settings.state & LVIS_STATEIMAGEMASK) >> 12);
	} else {
		*pValue = 1;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.mask & LVIF_TEXT) {
		if(properties.settings.pszText == LPSTR_TEXTCALLBACK) {
			*pValue = NULL;
		} else {
			*pValue = _bstr_t(properties.settings.pszText).Detach();
		}
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewItem::get_TileViewColumns(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	VariantClear(pValue);

	if((properties.settings.mask & LVIF_COLUMNS) == LVIF_COLUMNS || (properties.settings.mask & LVIF_COLFMT) == LVIF_COLFMT) {
		if(properties.settings.cColumns == I_COLUMNSCALLBACK) {
			// return 'Empty'
			VariantClear(pValue);
			hr = S_OK;
		} else if(properties.settings.cColumns > 0) {
			// return an array
			VariantClear(pValue);
			CComPtr<IRecordInfo> pRecordInfo = NULL;
			CLSID clsidTILEVIEWSUBITEM = {0};
			#ifdef UNICODE
				CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
			#else
				CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
			#endif
			pValue->parray = SafeArrayCreateVectorEx(VT_RECORD, 0, properties.settings.cColumns, pRecordInfo);
			pValue->vt = VT_ARRAY | VT_RECORD;

			TILEVIEWSUBITEM element = {0};
			for(LONG i = 0; i < static_cast<LONG>(properties.settings.cColumns); ++i) {
				element.ColumnIndex = properties.settings.puColumns[i];
				element.BeginNewColumn = BOOL2VARIANTBOOL(properties.settings.piColFmt[i] & LVCFMT_LINE_BREAK);
				element.FillRemainder = BOOL2VARIANTBOOL(properties.settings.piColFmt[i] & LVCFMT_FILL);
				element.WrapToMultipleLines = BOOL2VARIANTBOOL(properties.settings.piColFmt[i] & LVCFMT_WRAP);
				element.NoTitle = BOOL2VARIANTBOOL(properties.settings.piColFmt[i] & LVCFMT_NO_TITLE);
				ATLVERIFY(SUCCEEDED(SafeArrayPutElement(pValue->parray, &i, &element)));
			}
			hr = S_OK;
		} else {
			// return an empty array
			VariantClear(pValue);
			pValue->parray = NULL;
			pValue->vt = VT_ARRAY | VT_RECORD;
			hr = S_OK;
		}
	}
	return hr;
}