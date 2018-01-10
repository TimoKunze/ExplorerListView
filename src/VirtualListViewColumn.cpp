// VirtualListViewColumn.cpp: A wrapper for non-existing listview columns.

#include "stdafx.h"
#include "VirtualListViewColumn.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualListViewColumn::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualListViewColumn, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualListViewColumn::Properties::~Properties()
{
	if(headerSettings.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(headerSettings.pszText);
	}
	if(headerSettings.mask & HDI_FILTER) {
		if(headerSettings.type == HDFT_ISSTRING) {
			SECUREFREE(reinterpret_cast<LPHDTEXTFILTER>(headerSettings.pvFilter)->pszText);
		}
		SECUREFREE(headerSettings.pvFilter);
	}
	if(columnSettings.pszText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(columnSettings.pszText);
	}
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND VirtualListViewColumn::Properties::GetHeaderHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWndHeader(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualListViewColumn::LoadState(LPHDITEM pSourceHDITEM, LPLVCOLUMN pSourceLVCOLUMN)
{
	ATLASSERT_POINTER(pSourceHDITEM, HDITEM);
	#ifdef DEBUG
		if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
			ATLASSERT_POINTER(pSourceLVCOLUMN, LVCOLUMN);
		}
	#endif

	SECUREFREE(properties.headerSettings.pszText);
	properties.headerSettings = *pSourceHDITEM;
	if(properties.headerSettings.mask & HDI_TEXT) {
		// duplicate the column's text
		if(properties.headerSettings.pszText != LPSTR_TEXTCALLBACK) {
			properties.headerSettings.cchTextMax = lstrlen(pSourceHDITEM->pszText);
			properties.headerSettings.pszText = reinterpret_cast<LPTSTR>(malloc((properties.headerSettings.cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(properties.headerSettings.pszText, properties.headerSettings.cchTextMax + 1, pSourceHDITEM->pszText)));
		}
	}

	if(pSourceLVCOLUMN) {
		properties.columnSettings = *pSourceLVCOLUMN;
		// don't store the caption twice
		properties.columnSettings.pszText = NULL;
	}

	return S_OK;
}

HRESULT VirtualListViewColumn::SaveState(LPHDITEM pTargetHDITEM, LPLVCOLUMN pTargetLVCOLUMN)
{
	ATLASSERT_POINTER(pTargetHDITEM, HDITEM);
	#ifdef DEBUG
		if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
			ATLASSERT_POINTER(pTargetLVCOLUMN, LVCOLUMN);
		}
	#endif

	SECUREFREE(pTargetHDITEM->pszText);
	*pTargetHDITEM = properties.headerSettings;
	if(pTargetHDITEM->mask & HDI_TEXT) {
		// duplicate the column's text
		if(pTargetHDITEM->pszText != LPSTR_TEXTCALLBACK) {
			pTargetHDITEM->pszText = reinterpret_cast<LPTSTR>(malloc((pTargetHDITEM->cchTextMax + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(pTargetHDITEM->pszText, pTargetHDITEM->cchTextMax + 1, properties.headerSettings.pszText)));
		}
	}

	if(pTargetLVCOLUMN) {
		*pTargetLVCOLUMN = properties.columnSettings;
		// we didn't store the caption twice
		pTargetLVCOLUMN->pszText = pTargetHDITEM->pszText;
	}

	return S_OK;
}

void VirtualListViewColumn::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP VirtualListViewColumn::get_Alignment(AlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_FORMAT) {
		switch(properties.headerSettings.fmt & HDF_JUSTIFYMASK) {
			case HDF_LEFT:
				*pValue = alLeft;
				break;
			case HDF_CENTER:
				*pValue = alCenter;
				break;
			case HDF_RIGHT:
				*pValue = alRight;
				break;
		}
	} else {
		*pValue = alLeft;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_BitmapHandle(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	if(properties.headerSettings.mask & HDI_FORMAT) {
		if(properties.headerSettings.fmt & HDF_BITMAP) {
			if(properties.headerSettings.mask & HDI_BITMAP) {
				*pValue = HandleToLong(properties.headerSettings.hbm);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_Caption(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_TEXT) {
		if(properties.headerSettings.pszText == LPSTR_TEXTCALLBACK) {
			*pValue = NULL;
		} else {
			*pValue = _bstr_t(properties.headerSettings.pszText).Detach();
		}
	} else {
		*pValue = SysAllocString(L"");
	}
	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_STATE) {
		*pValue = BOOL2VARIANTBOOL(properties.headerSettings.state & HDIS_FOCUSED);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_ColumnData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_LPARAM) {
		*pValue = static_cast<LONG>(properties.headerSettings.lParam);
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_DefaultWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.columnSettings.mask & LVCF_DEFAULTWIDTH) {
		*pValue = properties.columnSettings.cxDefault;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_Filter(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	VariantClear(pValue);

	if(properties.headerSettings.mask & HDI_FILTER) {
		switch(properties.headerSettings.type) {
			case HDFT_ISSTRING:
				pValue->bstrVal = _bstr_t(reinterpret_cast<LPHDTEXTFILTER>(properties.headerSettings.pvFilter)->pszText).Detach();
				pValue->vt = VT_BSTR;
				break;
			case HDFT_ISNUMBER:
				pValue->intVal = *reinterpret_cast<PINT>(properties.headerSettings.pvFilter);
				pValue->vt = VT_INT;
				break;
			case HDFT_ISDATE: {
				SystemTimeToVariantTime(reinterpret_cast<LPSYSTEMTIME>(properties.headerSettings.pvFilter), &pValue->date);
				pValue->vt = VT_DATE;
				break;
			}
			case HDFT_HASNOVALUE:
				// nothing to do...
				break;
		}
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = -2;
	if(properties.headerSettings.mask & HDI_FORMAT) {
		if(properties.headerSettings.fmt & HDF_IMAGE) {
			if(properties.headerSettings.mask & HDI_IMAGE) {
				*pValue = properties.headerSettings.iImage;
			}
		}
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_ImageOnRight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_FORMAT) {
		*pValue = BOOL2VARIANTBOOL(properties.headerSettings.fmt & HDF_BITMAP_ON_RIGHT);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_MinimumWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.columnSettings.mask & LVCF_MINWIDTH) {
		*pValue = properties.columnSettings.cxMin;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_OwnerDrawn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_FORMAT) {
		*pValue = BOOL2VARIANTBOOL(properties.headerSettings.fmt & HDF_OWNERDRAW);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_Position(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_ORDER) {
		*pValue = properties.headerSettings.iOrder;
	} else {
		*pValue = 0;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_Resizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.columnSettings.mask & LVCF_FMT) {
		*pValue = BOOL2VARIANTBOOL((properties.columnSettings.fmt & LVCFMT_FIXED_WIDTH) == 0);
	} else {
		*pValue = VARIANT_TRUE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_ShowDropDownButton(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.columnSettings.mask & LVCF_FMT) {
		*pValue = BOOL2VARIANTBOOL(properties.columnSettings.fmt & LVCFMT_SPLITBUTTON);
	} else {
		*pValue = VARIANT_FALSE;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_SortArrow(SortArrowConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SortArrowConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_FORMAT) {
		switch(properties.headerSettings.fmt & (HDF_SORTDOWN | HDF_SORTUP)) {
			case HDF_SORTDOWN:
				*pValue = saDown;
				break;
			case HDF_SORTUP:
				*pValue = saUp;
				break;
			default:
				*pValue = saNone;
				break;
		}
	} else {
		*pValue = saNone;
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 0;
	if(properties.headerSettings.mask & HDI_FORMAT) {
		if(properties.headerSettings.fmt & HDF_CHECKBOX) {
			*pValue = ((properties.headerSettings.fmt & HDF_CHECKED) == HDF_CHECKED ? 2 : 1);
		}
	}

	return S_OK;
}

STDMETHODIMP VirtualListViewColumn::get_Width(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.headerSettings.mask & HDI_WIDTH) {
		*pValue = properties.headerSettings.cxy;
	} else {
		*pValue = 0;
	}

	return S_OK;
}