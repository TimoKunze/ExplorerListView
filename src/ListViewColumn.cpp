// ListViewColumn.cpp: A wrapper for existing listview columns.

#include "stdafx.h"
#include "ListViewColumn.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewColumn::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewColumn, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListViewColumn::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewColumn::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HWND ListViewColumn::Properties::GetHeaderHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWndHeader(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT ListViewColumn::SaveState(int columnIndex, LPHDITEM pTargetHDITEM, LPLVCOLUMN pTargetLVCOLUMN, HWND hWndHeader/* = NULL*/, HWND hWndLvw/* = NULL*/)
{
	if(!hWndHeader) {
		hWndHeader = properties.GetHeaderHWnd();
	}
	ATLASSERT(IsWindow(hWndHeader));
	if((columnIndex < 0) || (columnIndex >= SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTargetHDITEM, HDITEM);
	if(pTargetHDITEM == NULL) {
		return E_POINTER;
	}

	ZeroMemory(pTargetHDITEM, sizeof(HDITEM));
	pTargetHDITEM->cchTextMax = MAX_COLUMNCAPTIONLENGTH;
	pTargetHDITEM->pszText = reinterpret_cast<LPTSTR>(malloc((pTargetHDITEM->cchTextMax + 1) * sizeof(TCHAR)));
	pTargetHDITEM->mask = HDI_BITMAP | HDI_FORMAT | HDI_FILTER | HDI_HEIGHT | HDI_IMAGE | HDI_LPARAM | HDI_ORDER | HDI_STATE | HDI_TEXT | HDI_WIDTH;

	HRESULT hr = E_FAIL;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(pTargetHDITEM))) {
		hr = S_OK;
	}

	if(pTargetLVCOLUMN) {
		if(!hWndLvw) {
			hWndLvw = properties.GetExLvwHWnd();
		}
		ATLASSERT(IsWindow(hWndLvw));

		ZeroMemory(pTargetLVCOLUMN, sizeof(LVCOLUMN));
		pTargetLVCOLUMN->mask = LVCF_FMT | LVCF_MINWIDTH | LVCF_DEFAULTWIDTH | LVCF_IDEALWIDTH;
		SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(pTargetLVCOLUMN));
	}

	return hr;
}


void ListViewColumn::Attach(int columnIndex)
{
	properties.columnIndex = columnIndex;
}

void ListViewColumn::Detach(void)
{
	properties.columnIndex = -1;
}

HRESULT ListViewColumn::LoadState(LPHDITEM pSourceHDITEM, LPLVCOLUMN pSourceLVCOLUMN)
{
	ATLASSERT_POINTER(pSourceHDITEM, HDITEM);
	if(pSourceHDITEM == NULL) {
		return E_POINTER;
	}
	BOOL comctl32v610 = properties.pOwnerExLvw->IsComctl32Version610OrNewer();
	if(comctl32v610) {
		comctl32v610 = TRUE;
		ATLASSERT_POINTER(pSourceLVCOLUMN, LVCOLUMN);
		if(!pSourceLVCOLUMN) {
			return E_POINTER;
		}
	}

	// now apply the settings
	AlignmentConstants alignment = alLeft;
	if(pSourceHDITEM->mask & HDI_FORMAT) {
		switch(pSourceHDITEM->fmt & HDF_JUSTIFYMASK) {
			case HDF_LEFT:
				alignment = alLeft;
				break;
			case HDF_CENTER:
				alignment = alCenter;
				break;
			case HDF_RIGHT:
				alignment = alRight;
				break;
		}
		put_Alignment(alignment);
	}

	OLE_HANDLE h = 0;
	if(pSourceHDITEM->mask & HDI_BITMAP) {
		h = HandleToLong(pSourceHDITEM->hbm);
		put_BitmapHandle(h);
	}

	BSTR text = NULL;
	if(pSourceHDITEM->mask & HDI_TEXT) {
		if(pSourceHDITEM->pszText != LPSTR_TEXTCALLBACK) {
			text = _bstr_t(pSourceHDITEM->pszText).Detach();
		}
		put_Caption(text);
	}
	SysFreeString(text);

	VARIANT v;
	VariantInit(&v);
	if(pSourceHDITEM->mask & HDI_FILTER) {
		switch(pSourceHDITEM->type) {
			case HDFT_ISSTRING:
				v.bstrVal = _bstr_t(reinterpret_cast<LPHDTEXTFILTER>(pSourceHDITEM->pvFilter)->pszText).Detach();
				v.vt = VT_BSTR;
				break;
			case HDFT_ISNUMBER:
				v.intVal = *reinterpret_cast<PINT>(pSourceHDITEM->pvFilter);
				v.vt = VT_INT;
				break;
			case HDFT_ISDATE: {
				SystemTimeToVariantTime(reinterpret_cast<LPSYSTEMTIME>(pSourceHDITEM->pvFilter), &v.date);
				v.vt = VT_DATE;
				break;
			}
			case HDFT_HASNOVALUE:
				// nothing to do...
				break;
		}
		put_Filter(v);
	}
	VariantClear(&v);

	LONG l = 0;
	if(pSourceHDITEM->mask & HDI_LPARAM) {
		l = static_cast<LONG>(pSourceHDITEM->lParam);
		put_ColumnData(l);
	}
	l = 0;
	if(pSourceHDITEM->mask & HDI_IMAGE) {
		l = pSourceHDITEM->iImage;
		put_IconIndex(l);
	}
	l = 0;
	if(pSourceHDITEM->mask & HDI_ORDER) {
		l = pSourceHDITEM->iOrder;
		put_Position(l);
	}
	if(comctl32v610) {
		if(pSourceHDITEM->mask & HDI_FORMAT) {
			l = 0;
			if(pSourceHDITEM->fmt & HDF_CHECKBOX) {
				l = ((pSourceHDITEM->fmt & HDF_CHECKED) == HDF_CHECKED ? 2 : 1);
			}
			put_StateImageIndex(l);
		}
	}
	l = 0;
	if(pSourceHDITEM->mask & HDI_WIDTH) {
		l = pSourceHDITEM->cxy;
		put_Width(l);
	}

	VARIANT_BOOL b = VARIANT_FALSE;
	if(pSourceHDITEM->mask & HDI_FORMAT) {
		b = BOOL2VARIANTBOOL(pSourceHDITEM->fmt & HDF_BITMAP_ON_RIGHT);
		put_ImageOnRight(b);
		b = BOOL2VARIANTBOOL(pSourceHDITEM->fmt & HDF_OWNERDRAW);
		put_OwnerDrawn(b);
	}
	if(pSourceLVCOLUMN->mask & LVCF_FMT) {
		if(comctl32v610) {
			b = BOOL2VARIANTBOOL((pSourceLVCOLUMN->fmt & LVCFMT_FIXED_WIDTH) == 0);
			put_Resizable(b);
			b = BOOL2VARIANTBOOL(pSourceLVCOLUMN->fmt & LVCFMT_SPLITBUTTON);
			put_ShowDropDownButton(b);
		}
	}

	SortArrowConstants sortArrow = saNone;
	if(pSourceHDITEM->mask & HDI_FORMAT) {
		switch(pSourceHDITEM->fmt & (HDF_SORTDOWN | HDF_SORTUP)) {
			case HDF_SORTDOWN:
				sortArrow = saDown;
				break;
			case HDF_SORTUP:
				sortArrow = saUp;
				break;
			default:
				sortArrow = saNone;
				break;
		}
		put_SortArrow(sortArrow);
	}

	return S_OK;
}

HRESULT ListViewColumn::LoadState(VirtualListViewColumn* pSource)
{
	ATLASSUME(pSource);
	if(pSource == NULL) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	// now apply the settings
	AlignmentConstants alignment = alLeft;
	pSource->get_Alignment(&alignment);
	put_Alignment(alignment);

	OLE_HANDLE h = 0;
	pSource->get_BitmapHandle(&h);
	put_BitmapHandle(h);

	BSTR text;
	pSource->get_Caption(&text);
	put_Caption(text);
	SysFreeString(text);

	VARIANT v;
	VariantInit(&v);
	pSource->get_Filter(&v);
	put_Filter(v);
	VariantClear(&v);

	LONG l = 0;
	pSource->get_ColumnData(&l);
	put_ColumnData(l);
	l = 0;
	pSource->get_IconIndex(&l);
	put_IconIndex(l);
	l = 0;
	if(pSource->get_Position(&l) != S_OK) {
		l = static_cast<LONG>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
	}
	put_Position(l);
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		l = 0;
		pSource->get_StateImageIndex(&l);
		put_StateImageIndex(l);
	}
	l = 120;
	pSource->get_Width(&l);
	put_Width(l);

	VARIANT_BOOL b = VARIANT_FALSE;
	pSource->get_ImageOnRight(&b);
	put_ImageOnRight(b);
	b = VARIANT_FALSE;
	pSource->get_OwnerDrawn(&b);
	put_OwnerDrawn(b);
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		b = VARIANT_TRUE;
		pSource->get_Resizable(&b);
		put_Resizable(b);
		b = VARIANT_FALSE;
		pSource->get_ShowDropDownButton(&b);
		put_ShowDropDownButton(b);
	}

	SortArrowConstants sortArrow = saNone;
	pSource->get_SortArrow(&sortArrow);
	put_SortArrow(sortArrow);

	return S_OK;
}

HRESULT ListViewColumn::SaveState(LPHDITEM pTargetHDITEM, LPLVCOLUMN pTargetLVCOLUMN, HWND hWndHeader/* = NULL*/, HWND hWndLvw/* = NULL*/)
{
	return SaveState(properties.columnIndex, pTargetHDITEM, pTargetLVCOLUMN, hWndHeader, hWndLvw);
}

HRESULT ListViewColumn::SaveState(VirtualListViewColumn* pTarget)
{
	ATLASSUME(pTarget);
	if(pTarget == NULL) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerExLvw);
	HDITEM columnHeader = {0};
	LVCOLUMN column = {0};
	HRESULT hr = SaveState(&columnHeader, &column);
	pTarget->LoadState(&columnHeader, &column);
	SECUREFREE(columnHeader.pszText);
	if(columnHeader.pvFilter) {
		if(columnHeader.type == HDFT_ISSTRING) {
			SECUREFREE(reinterpret_cast<LPHDTEXTFILTER>(columnHeader.pvFilter)->pszText);
		}
		SECUREFREE(columnHeader.pvFilter);
	}

	return hr;
}

void ListViewColumn::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewColumn::get_Alignment(AlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVCOLUMN column = {0};
	column.mask = LVCF_FMT;
	if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		switch(column.fmt & LVCFMT_JUSTIFYMASK) {
			case LVCFMT_LEFT:
				*pValue = alLeft;
				break;
			case LVCFMT_CENTER:
				*pValue = alCenter;
				break;
			case LVCFMT_RIGHT:
				*pValue = alRight;
				break;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_Alignment(AlignmentConstants newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVCOLUMN column = {0};
	column.mask = LVCF_FMT;
	SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

	switch(newValue) {
		case alLeft:
			column.fmt &= ~(LVCFMT_CENTER | LVCFMT_RIGHT);
			column.fmt |= LVCFMT_LEFT;
			break;
		case alCenter:
			column.fmt &= ~(LVCFMT_LEFT | LVCFMT_RIGHT);
			column.fmt |= LVCFMT_CENTER;
			break;
		case alRight:
			column.fmt &= ~(LVCFMT_LEFT | LVCFMT_CENTER);
			column.fmt |= LVCFMT_RIGHT;
			break;
	}
	if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		InvalidateRect(hWndLvw, NULL, TRUE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_BitmapHandle(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT | HDI_BITMAP;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		if(column.fmt & HDF_BITMAP) {
			*pValue = HandleToLong(column.hbm);
		} else {
			*pValue = 0;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_BitmapHandle(OLE_HANDLE newValue)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

	BOOL freeOnFailure = FALSE;
	if(newValue == -1) {
		// use Windows' default down arrow
		newValue = 0;
		HMODULE hMod = LoadLibrary(TEXT("shell32.dll"));
		if(hMod) {
			HANDLE hBMP = LoadImage(hMod, MAKEINTRESOURCE(IDB_SHELL32_DOWNARROW), IMAGE_BITMAP, 14, 14, LR_LOADMAP3DCOLORS);
			if(hBMP) {
				newValue = HandleToLong(hBMP);
				freeOnFailure = TRUE;
			}
			FreeLibrary(hMod);
		}
	} else if(newValue == -2) {
		// use Windows' default up arrow
		newValue = 0;
		HMODULE hMod = LoadLibrary(TEXT("shell32.dll"));
		if(hMod) {
			HANDLE hBMP = LoadImage(hMod, MAKEINTRESOURCE(IDB_SHELL32_UPARROW), IMAGE_BITMAP, 14, 14, LR_LOADMAP3DCOLORS);
			if(hBMP) {
				newValue = HandleToLong(hBMP);
				freeOnFailure = TRUE;
			}
			FreeLibrary(hMod);
		}
	}

	if(newValue == 0) {
		column.fmt &= ~HDF_BITMAP;
	} else {
		column.mask |= HDI_BITMAP;
		column.fmt |= HDF_BITMAP;
		column.hbm = static_cast<HBITMAP>(LongToHandle(newValue));
	}
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return S_OK;
	}
	if(freeOnFailure) {
		DeleteObject(column.hbm);
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Caption(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.cchTextMax = MAX_COLUMNCAPTIONLENGTH;
	column.pszText = reinterpret_cast<LPTSTR>(malloc((column.cchTextMax + 1) * sizeof(TCHAR)));
	LPTSTR pBuffer = column.pszText;
	column.mask = HDI_TEXT;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		*pValue = _bstr_t(column.pszText).Detach();
		hr = S_OK;
	}
	// according to Windows SDK, the header may decide to store the string into another buffer
	if(column.pszText != pBuffer) {
		SECUREFREE(pBuffer);
	}
	SECUREFREE(column.pszText);
	return hr;
}

STDMETHODIMP ListViewColumn::put_Caption(BSTR newValue)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HRESULT hr = E_FAIL;
	HDITEM column = {0};
	column.mask = HDI_TEXT;
	COLE2T converter(newValue);
	if(newValue) {
		column.pszText = converter;
	} else {
		column.pszText = LPSTR_TEXTCALLBACK;
	}
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewColumn::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndHeader = properties.GetHeaderHWnd();
		ATLASSERT(IsWindow(hWndHeader));

		HDITEM column = {0};
		column.mask = HDI_STATE;
		if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			*pValue = BOOL2VARIANTBOOL(column.state & HDIS_FOCUSED);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_ColumnData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_LPARAM;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		*pValue = static_cast<LONG>(column.lParam);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_ColumnData(LONG newValue)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_LPARAM;
	column.lParam = newValue;
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_DefaultWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_DEFAULTWIDTH;
		if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			*pValue = column.cxDefault;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_DefaultWidth(OLE_XSIZE_PIXELS newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.cxDefault = static_cast<int>(newValue);
		column.mask = LVCF_DEFAULTWIDTH;
		if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Filter(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	//if(properties.pOwnerExLvw->IsComctl32Version580OrNewer()) {
		HWND hWndHeader = properties.GetHeaderHWnd();
		ATLASSERT(IsWindow(hWndHeader));

		HDITEM column = {0};
		column.mask = HDI_FILTER;
		// the first call gives us the data type (integer/string)
		if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			switch(column.type) {
				case HDFT_ISSTRING: {
					HDTEXTFILTER filter = {0};
					filter.cchTextMax = 1024;
					filter.pszText = new TCHAR[filter.cchTextMax + 1];
					column.pvFilter = &filter;
					if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
						VariantClear(pValue);
						pValue->bstrVal = _bstr_t(filter.pszText).Detach();
						pValue->vt = VT_BSTR;

						hr = S_OK;
					}
					delete[] filter.pszText;
					break;
				}
				case HDFT_ISNUMBER:
					column.pvFilter = new int;
					if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
						VariantClear(pValue);
						pValue->intVal = *reinterpret_cast<PINT>(column.pvFilter);
						pValue->vt = VT_INT;

						hr = S_OK;
					}
					delete column.pvFilter;
					break;
				case HDFT_ISDATE:
					column.pvFilter = new SYSTEMTIME;
					if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
						VariantClear(pValue);
						SystemTimeToVariantTime(reinterpret_cast<LPSYSTEMTIME>(column.pvFilter), &pValue->date);
						pValue->vt = VT_DATE;

						hr = S_OK;
					}
					delete column.pvFilter;
					break;
				case HDFT_HASNOVALUE:
					VariantClear(pValue);
					hr = S_OK;
					break;
			}
		}
	//}
	return hr;
}

STDMETHODIMP ListViewColumn::put_Filter(VARIANT newValue)
{
	//if(properties.pOwnerExLvw->IsComctl32Version580OrNewer()) {
		HWND hWndHeader = properties.GetHeaderHWnd();
		ATLASSERT(IsWindow(hWndHeader));

		if(newValue.vt == VT_EMPTY) {
			if(SendMessage(hWndHeader, HDM_CLEARFILTER, properties.columnIndex, 0)) {
				return S_OK;
			}
		} else {
			HDITEM column = {0};
			HDTEXTFILTER filter = {0};
			BSTR tmp = NULL;
			if(newValue.vt == VT_BSTR) {
				tmp = newValue.bstrVal;
			}
			COLE2T converter(tmp);
			SYSTEMTIME stDummy = {0};
			int intDummy = 0;
			if(newValue.vt == VT_BSTR) {
				filter.pszText = converter;
				filter.cchTextMax = lstrlen(filter.pszText);
				column.pvFilter = &filter,
				column.type = HDFT_ISSTRING;
			} else if(newValue.vt == VT_DATE && properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
				VariantTimeToSystemTime(newValue.date, &stDummy);
				column.pvFilter = &stDummy;
				column.type = HDFT_ISDATE;
			} else {
				VARIANT v;
				VariantInit(&v);
				if(FAILED(VariantChangeType(&v, &newValue, 0, VT_INT))) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				intDummy = v.intVal;
				column.pvFilter = &intDummy;
				column.type = HDFT_ISNUMBER;
			}
			column.mask = HDI_FILTER;
			if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
				return S_OK;
			}
		}
	//}

	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Height(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	CRect rc;
	if(SendMessage(hWndHeader, HDM_GETITEMRECT, properties.columnIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.Height();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT | HDI_IMAGE;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		if(column.fmt & HDF_IMAGE) {
			*pValue = column.iImage;
		} else {
			*pValue = -2;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_IconIndex(LONG newValue)
{
	if(newValue < -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

	if(newValue == -2) {
		column.fmt &= ~HDF_IMAGE;
	} else {
		column.mask |= HDI_IMAGE;
		column.fmt |= HDF_IMAGE;
		column.iImage = newValue;
	}
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.columnIndex < 0) {
		return E_FAIL;
	}

	if(properties.pOwnerExLvw) {
		*pValue = properties.pOwnerExLvw->ColumnIndexToID(properties.columnIndex);
		if(*pValue != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_IdealWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_IDEALWIDTH;
		if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			*pValue = column.cxIdeal;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_ImageOnRight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		*pValue = BOOL2VARIANTBOOL(column.fmt & HDF_BITMAP_ON_RIGHT);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_ImageOnRight(VARIANT_BOOL newValue)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

	if(newValue == VARIANT_FALSE) {
		column.fmt &= ~HDF_BITMAP_ON_RIGHT;
	} else {
		column.fmt |= HDF_BITMAP_ON_RIGHT;
	}
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.columnIndex;
	return S_OK;
}

STDMETHODIMP ListViewColumn::get_Left(OLE_XPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	RECT rc = {0};
	if(SendMessage(hWndHeader, HDM_GETITEMRECT, properties.columnIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.left;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Locale(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.pOwnerExLvw->GetColumnLocale(properties.columnIndex);
	if(*pValue == -1) {
		*pValue = GetThreadLocale();
	}
	return S_OK;
}

STDMETHODIMP ListViewColumn::put_Locale(LONG newValue)
{
	properties.pOwnerExLvw->SetColumnLocale(properties.columnIndex, newValue);
	return S_OK;
}

STDMETHODIMP ListViewColumn::get_MinimumWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_MINWIDTH;
		if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			*pValue = column.cxMin;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_MinimumWidth(OLE_XSIZE_PIXELS newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.cxMin = newValue;
		column.mask = LVCF_MINWIDTH;
		if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_OwnerDrawn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		*pValue = BOOL2VARIANTBOOL(column.fmt & HDF_OWNERDRAW);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_OwnerDrawn(VARIANT_BOOL newValue)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

	if(newValue == VARIANT_FALSE) {
		column.fmt &= ~HDF_OWNERDRAW;
	} else {
		column.fmt |= HDF_OWNERDRAW;
	}
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Position(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVCOLUMN column = {0};
	column.mask = LVCF_ORDER;
	if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		*pValue = column.iOrder;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_Position(LONG newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVCOLUMN column = {0};
	column.iOrder = newValue;
	column.mask = LVCF_ORDER;
	if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		InvalidateRect(hWndLvw, NULL, TRUE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Resizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_FMT;
		if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			*pValue = BOOL2VARIANTBOOL((column.fmt & LVCFMT_FIXED_WIDTH) == 0);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_Resizable(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_FMT;
		SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

		if(newValue == VARIANT_FALSE) {
			column.fmt |= LVCFMT_FIXED_WIDTH;
		} else {
			column.fmt &= ~LVCFMT_FIXED_WIDTH;
		}
		if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(RunTimeHelper::IsCommCtrl6()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		*pValue = BOOL2VARIANTBOOL(SendMessage(hWndLvw, LVM_GETSELECTEDCOLUMN, 0, 0) == properties.columnIndex);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	STDMETHODIMP ListViewColumn::get_ShellListViewColumnObject(IDispatch** ppColumn)
	{
		ATLASSERT_POINTER(ppColumn, LPDISPATCH);
		if(!ppColumn) {
			return E_POINTER;
		}

		if(properties.columnIndex < 0) {
			return E_FAIL;
		}

		if(properties.pOwnerExLvw) {
			if(properties.pOwnerExLvw->shellBrowserInterface.pInternalMessageListener) {
				LONG columnID = properties.pOwnerExLvw->ColumnIndexToID(properties.columnIndex);
				if(columnID != -1) {
					properties.pOwnerExLvw->shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_GETSHLISTVIEWCOLUMN, columnID, reinterpret_cast<LPARAM>(ppColumn));
					return S_OK;
				}
			}
		}
		return E_FAIL;
	}
#endif

STDMETHODIMP ListViewColumn::get_ShowDropDownButton(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_FMT;
		if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			*pValue = BOOL2VARIANTBOOL(column.fmt & LVCFMT_SPLITBUTTON);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_ShowDropDownButton(VARIANT_BOOL newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVCOLUMN column = {0};
		column.mask = LVCF_FMT;
		SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

		if(newValue == VARIANT_FALSE) {
			column.fmt &= ~LVCFMT_SPLITBUTTON;
		} else {
			column.fmt |= LVCFMT_SPLITBUTTON;
		}
		if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_SortArrow(SortArrowConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SortArrowConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		switch(column.fmt & (HDF_SORTDOWN | HDF_SORTUP)) {
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
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_SortArrow(SortArrowConstants newValue)
{
	if((newValue != saNone) && !RunTimeHelper::IsCommCtrl6()) {
		return E_FAIL;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

	switch(newValue) {
		case saDown:
			column.fmt &= ~HDF_SORTUP;
			column.fmt |= HDF_SORTDOWN;
			break;
		case saUp:
			column.fmt &= ~HDF_SORTDOWN;
			column.fmt |= HDF_SORTUP;
			break;
		case saNone:
			column.fmt &= ~(HDF_SORTDOWN | HDF_SORTUP);
			break;
	}
	if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		HWND hWndHeader = properties.GetHeaderHWnd();
		ATLASSERT(IsWindow(hWndHeader));

		HDITEM column = {0};
		column.mask = HDI_FORMAT;
		if(SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			if(column.fmt & HDF_CHECKBOX) {
				*pValue = ((column.fmt & HDF_CHECKED) == HDF_CHECKED ? 2 : 1);
			} else {
				*pValue = 0;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_StateImageIndex(LONG newValue)
{
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		if(newValue < 0) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}

		HWND hWndHeader = properties.GetHeaderHWnd();
		ATLASSERT(IsWindow(hWndHeader));

		HDITEM column = {0};
		column.mask = HDI_FORMAT;
		SendMessage(hWndHeader, HDM_GETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column));

		if(newValue == 0) {
			column.fmt &= ~(HDF_CHECKBOX | HDF_CHECKED);
		} else {
			column.fmt |= HDF_CHECKBOX;
			switch(newValue) {
				case 1:
					column.fmt &= ~HDF_CHECKED;
					break;
				case 2:
					column.fmt |= HDF_CHECKED;
					break;
			}
		}
		if(SendMessage(hWndHeader, HDM_SETITEM, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.pOwnerExLvw->GetColumnTextParsingFlags(properties.columnIndex, parsingFunction);
	return S_OK;
}

STDMETHODIMP ListViewColumn::put_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG newValue)
{
	properties.pOwnerExLvw->SetColumnTextParsingFlags(properties.columnIndex, parsingFunction, newValue);
	return S_OK;
}

STDMETHODIMP ListViewColumn::get_Top(OLE_YPOS_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YPOS_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	RECT rc = {0};
	if(SendMessage(hWndHeader, HDM_GETITEMRECT, properties.columnIndex, reinterpret_cast<LPARAM>(&rc))) {
		*pValue = rc.top;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::get_Width(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVCOLUMN column = {0};
	column.mask = LVCF_WIDTH;
	if(SendMessage(hWndLvw, LVM_GETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
		*pValue = column.cx;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::put_Width(OLE_XSIZE_PIXELS newValue)
{
	ATLASSERT(newValue >= LVSCW_AUTOSIZE_USEHEADER);
	if(newValue < LVSCW_AUTOSIZE_USEHEADER) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(newValue < 0) {
		if(SendMessage(hWndLvw, LVM_SETCOLUMNWIDTH, 0, MAKELPARAM(newValue, 0))) {
			return S_OK;
		}
	} else {
		LVCOLUMN column = {0};
		column.cx = newValue;
		column.mask = LVCF_WIDTH;
		if(SendMessage(hWndLvw, LVM_SETCOLUMN, properties.columnIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
	}
	return E_FAIL;
}


STDMETHODIMP ListViewColumn::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(phImageList == NULL) {
		return E_POINTER;
	}

	ATLASSUME(properties.pOwnerExLvw);

	// ask the native listview for a drag image
	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerExLvw->CreateLegacyHeaderDragImage(properties.columnIndex, &upperLeftPoint, NULL));

	if(*phImageList) {
		if(pXUpperLeft) {
			*pXUpperLeft = upperLeftPoint.x;
		}
		if(pYUpperLeft) {
			*pYUpperLeft = upperLeftPoint.y;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::GetDropDownRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		RECT dropDownRectangle = {0};
		if(SendMessage(hWndHeader, HDM_GETITEMDROPDOWNRECT, properties.columnIndex, reinterpret_cast<LPARAM>(&dropDownRectangle))) {
			if(pLeft) {
				*pLeft = dropDownRectangle.left;
			}
			if(pTop) {
				*pTop = dropDownRectangle.top;
			}
			if(pRight) {
				*pRight = dropDownRectangle.right;
			}
			if(pBottom) {
				*pBottom = dropDownRectangle.bottom;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewColumn::StartFilterEditing(VARIANT_BOOL applyCurrentFilter/* = VARIANT_TRUE*/)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	//if(properties.pOwnerExLvw->IsComctl32Version580OrNewer()) {
		if(SendMessage(hWndHeader, HDM_EDITFILTER, properties.columnIndex, !VARIANTBOOL2BOOL(applyCurrentFilter))) {
			return S_OK;
		}
	//}
	return E_FAIL;
}