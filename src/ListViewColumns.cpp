// ListViewColumns.cpp: Manages a collection of ListViewColumn objects

#include "stdafx.h"
#include "ListViewColumns.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewColumns::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewColumns, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListViewColumns::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(ppEnumerator == NULL) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListViewColumns>* pLvwColumnsObj = NULL;
	CComObject<ListViewColumns>::CreateInstance(&pLvwColumnsObj);
	pLvwColumnsObj->AddRef();

	// clone all settings
	pLvwColumnsObj->properties = properties;

	pLvwColumnsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLvwColumnsObj->Release();
	return S_OK;
}

STDMETHODIMP ListViewColumns::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(pItems == NULL) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		if(properties.nextEnumeratedColumn >= SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0)) {
			properties.nextEnumeratedColumn = 0;
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitColumn(properties.nextEnumeratedColumn, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal), FALSE);
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedColumn;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListViewColumns::Reset(void)
{
	properties.nextEnumeratedColumn = 0;
	return S_OK;
}

STDMETHODIMP ListViewColumns::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedColumn += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


ListViewColumns::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewColumns::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HWND ListViewColumns::Properties::GetHeaderHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWndHeader(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListViewColumns::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewColumns::get_Item(LONG columnIdentifier, ColumnIdentifierTypeConstants columnIdentifierType/* = citIndex*/, IListViewColumn** ppColumn/* = NULL*/)
{
	ATLASSERT_POINTER(ppColumn, IListViewColumn*);
	if(ppColumn == NULL) {
		return E_POINTER;
	}

	// get the column's index
	int columnIndex = -1;
	switch(columnIdentifierType) {
		case citIndex:
			if(columnIdentifier >= 0) {
				HWND hWndHeader = properties.GetHeaderHWnd();
				ATLASSERT(IsWindow(hWndHeader));

				int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
				if(columnCount == -1) {
					return E_FAIL;
				}
				if(columnIdentifier < columnCount) {
					columnIndex = columnIdentifier;
				}
			}
			break;
		case citPosition:
			if(columnIdentifier >= 0) {
				HWND hWndHeader = properties.GetHeaderHWnd();
				ATLASSERT(IsWindow(hWndHeader));

				int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
				if(columnCount == -1) {
					return E_FAIL;
				}
				if(columnIdentifier < columnCount) {
					columnIndex = static_cast<int>(SendMessage(hWndHeader, HDM_ORDERTOINDEX, columnIdentifier, 0));
				}
			}
			break;
		case citID:
			if(properties.pOwnerExLvw) {
				columnIndex = properties.pOwnerExLvw->IDToColumnIndex(columnIdentifier);
			}
			break;
	}

	if(columnIndex != -1) {
		ClassFactory::InitColumn(columnIndex, properties.pOwnerExLvw, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppColumn), FALSE);
		return S_OK;
	}

	return DISP_E_BADINDEX;
}

STDMETHODIMP ListViewColumns::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(ppEnumerator == NULL) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}

STDMETHODIMP ListViewColumns::get_Positions(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	VariantClear(pValue);

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	int columns = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
	if(columns == -1) {
		return E_FAIL;
	}
	if(columns == 0) {
		pValue->parray = NULL;
		pValue->vt = VT_ARRAY | VT_I4;
		return S_OK;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));
	PINT pBuffer = new int[columns];

	if(SendMessage(hWndLvw, LVM_GETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pBuffer))) {
		CComSafeArray<LONG> positions;
		positions.Create(columns);
		for(int i = 0; i < columns; ++i) {
			positions.SetAt(i, pBuffer[i]);
		}
		delete[] pBuffer;

		pValue->parray = SafeArrayCreateVectorEx(VT_I4, 0, columns, NULL);
		positions.CopyTo(&pValue->parray);
		pValue->vt = VT_ARRAY | VT_I4;
		return S_OK;
	}

	delete[] pBuffer;
	return E_FAIL;
}

STDMETHODIMP ListViewColumns::put_Positions(VARIANT newValue)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	int columns = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
	if(columns <= 0) {
		return E_FAIL;
	}

	PINT pPositions = NULL;
	if(newValue.vt & VT_ARRAY) {
		// an array
		LONG l = 0;
		SafeArrayGetLBound(newValue.parray, 1, &l);
		LONG u = 0;
		SafeArrayGetUBound(newValue.parray, 1, &u);
		if(u < l) {
			return E_INVALIDARG;
		}
		LONG arraySize = u - l + 1;
		if(columns != arraySize) {
			return E_INVALIDARG;
		}

		pPositions = new int[arraySize];

		VARTYPE vt = 0;
		SafeArrayGetVartype(newValue.parray, &vt);
		for(LONG i = l; i <= u; ++i) {
			if(vt == VT_VARIANT) {
				VARIANT v;
				SafeArrayGetElement(newValue.parray, &i, &v);
				if(SUCCEEDED(VariantChangeType(&v, &v, 0, VT_INT))) {
					pPositions[i - l] = v.intVal;
				} else {
					VariantClear(&v);
					delete[] pPositions;
					return E_INVALIDARG;
				}
				VariantClear(&v);
			} else if(vt == VT_I1 || vt == VT_UI1 || vt == VT_I2 || vt == VT_UI2 || vt == VT_I4 || vt == VT_UI4 || vt == VT_INT || vt == VT_UINT) {
				SafeArrayGetElement(newValue.parray, &i, &pPositions[i - l]);
			} else {
				delete[] pPositions;
				return E_INVALIDARG;
			}
		}
	} else {
		// a single value
		if(columns != 1) {
			return E_INVALIDARG;
		}

		pPositions = new int[1];
		VARIANT v;
		VariantInit(&v);
		if(SUCCEEDED(VariantChangeType(&v, &newValue, 0, VT_INT))) {
			pPositions[0] = v.intVal;
		} else {
			VariantClear(&v);
			delete[] pPositions;
			return E_INVALIDARG;
		}
		VariantClear(&v);
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(SendMessage(hWndLvw, LVM_SETCOLUMNORDERARRAY, columns, reinterpret_cast<LPARAM>(pPositions))) {
		InvalidateRect(hWndLvw, NULL, TRUE);
		delete[] pPositions;
		return S_OK;
	}
	delete[] pPositions;
	return E_FAIL;
}

STDMETHODIMP ListViewColumns::get_PositionsString(BSTR stringDelimiter, BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	VARIANT positionArray;
	VariantInit(&positionArray);
	HRESULT hr = get_Positions(&positionArray);
	if(SUCCEEDED(hr)) {
		if(positionArray.vt == (VT_ARRAY | VT_I4) && positionArray.parray) {
			// an array
			CAtlString positionsString;
			LONG l = 0;
			SafeArrayGetLBound(positionArray.parray, 1, &l);
			LONG u = 0;
			SafeArrayGetUBound(positionArray.parray, 1, &u);

			int position = 0;
			for(LONG i = l; i <= u; ++i) {
				SafeArrayGetElement(positionArray.parray, &i, &position);
				positionsString.AppendFormat(TEXT("%i"), position);
				if(i < u) {
					positionsString += stringDelimiter;
				}
			}
			*pValue = positionsString.AllocSysString();
		} else {
			ATLASSERT(FALSE && "get_Positions has been changed in an incompatible way");
			hr = E_FAIL;
		}
	}
	VariantClear(&positionArray);
	return hr;
}

STDMETHODIMP ListViewColumns::put_PositionsString(BSTR stringDelimiter, BSTR newValue)
{
	CAtlString delimiter(stringDelimiter);
	if(delimiter.Trim().GetLength() == 0) {
		return E_INVALIDARG;
	}

	CAtlString positionsString(newValue);
	// trim delimiters at start and end
	while(positionsString.Left(delimiter.GetLength()) == delimiter) {
		positionsString.Delete(0, delimiter.GetLength());
	}
	while(positionsString.Right(delimiter.GetLength()) == delimiter) {
		positionsString.Delete(positionsString.GetLength() - delimiter.GetLength(), delimiter.GetLength());
	}

	int numberOfElements = 0;
	if(positionsString.Trim().GetLength() > 0) {
		numberOfElements++;
		int p = 0;
		int start = 0;
		do {
			p = positionsString.Find(delimiter, start);
			if(p >= 0) {
				numberOfElements++;
				start = p + delimiter.GetLength();
			}
		} while(p >= 0);
	}

	VARIANT positionArray;
	VariantInit(&positionArray);
	if(numberOfElements == 0) {
		positionArray.parray = NULL;
		positionArray.vt = VT_ARRAY | VT_I4;
	} else {
		CComSafeArray<LONG> positions;
		positions.Create(numberOfElements);
		int start = 0;
		int element = 0;
		for(int i = 0; i < numberOfElements; ++i) {
			int p = positionsString.Find(delimiter, start);
			if(p >= 0) {
				if(p - start > 0) {
					element = _ttoi(positionsString.Mid(start, p - start));
				} else {
					element = -1;
				}
				positions.SetAt(i, element);
				start = p + delimiter.GetLength();
			} else {
				element = _ttoi(positionsString.Mid(start));
				positions.SetAt(i, element);
			}
		}

		positionArray.parray = SafeArrayCreateVectorEx(VT_I4, 0, numberOfElements, NULL);
		positions.CopyTo(&positionArray.parray);
		positionArray.vt = VT_ARRAY | VT_I4;
	}
	
	HRESULT hr = put_Positions(positionArray);
	VariantClear(&positionArray);
	return hr;
}


STDMETHODIMP ListViewColumns::Add(BSTR columnCaption, LONG insertAt/* = -1*/, LONG columnWidth/* = 120*/, LONG minimumWidth/* = 0*/, AlignmentConstants alignment/* = alLeft*/, LONG columnData/* = 0*/, LONG stateImageIndex/* = 1*/, VARIANT_BOOL resizable/* = VARIANT_TRUE*/, VARIANT_BOOL showDropDownButton/* = VARIANT_FALSE*/, VARIANT_BOOL ownerDrawn/* = VARIANT_FALSE*/, IListViewColumn** ppAddedColumn/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedColumn, IListViewColumn*);
	if(!ppAddedColumn) {
		return E_POINTER;
	}
	ATLASSERT(columnWidth >= LVSCW_AUTOSIZE_USEHEADER);
	if(columnWidth < -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	ATLASSERT(stateImageIndex >= 0);
	if(stateImageIndex < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	int columns = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
	if(insertAt == -1) {
		insertAt = columns;
	}
	ATLASSERT(insertAt >= 0 && insertAt <= columns);
	if(insertAt < 0 || insertAt > columns) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVCOLUMN column = {0};
	column.mask = LVCF_FMT | LVCF_TEXT;
	if(columnWidth >= 0) {
		column.mask |= LVCF_WIDTH;
		column.cx = columnWidth;
	}
	switch(alignment) {
		case alLeft:
			column.fmt = LVCFMT_LEFT;
			break;
		case alCenter:
			column.fmt = LVCFMT_CENTER;
			break;
		case alRight:
			column.fmt = LVCFMT_RIGHT;
			break;
	}
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		column.mask |= LVCF_MINWIDTH;
		column.cxMin = minimumWidth;
		if(resizable == VARIANT_FALSE) {
			column.fmt |= LVCFMT_FIXED_WIDTH;
		}
		if(showDropDownButton != VARIANT_FALSE) {
			column.fmt |= LVCFMT_SPLITBUTTON;
		}
	}

	// finally insert the column
	*ppAddedColumn = NULL;
	int insertedColumn = -1;
	#ifndef UNICODE
		COLE2T converter(columnCaption);
	#endif
	if(columnCaption) {
		#ifdef UNICODE
			column.pszText = OLE2W(columnCaption);
		#else
			column.pszText = converter;
		#endif
	} else {
		column.pszText = LPSTR_TEXTCALLBACK;
	}

	insertedColumn = static_cast<int>(SendMessage(hWndLvw, LVM_INSERTCOLUMN, insertAt, reinterpret_cast<LPARAM>(&column)));
	if(insertedColumn != -1) {
		BOOL postProcess = FALSE;
		if(CWindow(hWndHeader).GetExStyle() & WS_EX_RTLREADING) {
			postProcess = TRUE;
		}
		if(ownerDrawn != VARIANT_FALSE) {
			postProcess = TRUE;
		}

		if(postProcess) {
			HDITEM col = {0};
			col.mask = HDI_FORMAT;
			SendMessage(hWndHeader, HDM_GETITEM, insertedColumn, reinterpret_cast<LPARAM>(&col));
			if(CWindow(hWndHeader).GetExStyle() & WS_EX_RTLREADING) {
				col.fmt |= HDF_RTLREADING;
			}
			if(ownerDrawn != VARIANT_FALSE) {
				col.fmt |= HDF_OWNERDRAW;
			}
			SendMessage(hWndHeader, HDM_SETITEM, insertedColumn, reinterpret_cast<LPARAM>(&col));
		}
		if(insertAt == 0 && properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
			if(resizable == VARIANT_FALSE || showDropDownButton != VARIANT_FALSE) {
				/* HACK: Inserting a column with the LVCFMT_FIXED_WIDTH or LVCFMT_SPLITBUTTON flag at position 0
								 doesn't work, so set this flag after insertion. */
				LVCOLUMN col = {0};
				col.mask = LVCF_FMT;
				SendMessage(hWndLvw, LVM_GETCOLUMN, insertedColumn, reinterpret_cast<LPARAM>(&col));
				if(resizable == VARIANT_FALSE) {
					col.fmt |= LVCFMT_FIXED_WIDTH;
				}
				if(showDropDownButton != VARIANT_FALSE) {
					col.fmt |= LVCFMT_SPLITBUTTON;
				}
				SendMessage(hWndLvw, LVM_SETCOLUMN, insertedColumn, reinterpret_cast<LPARAM>(&col));
			}
		}

		ClassFactory::InitColumn(insertedColumn, properties.pOwnerExLvw, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppAddedColumn));
		if(columnData != 0) {
			// to set the column data, we have to work directly on the header, so do it _after_ insertion
			(*ppAddedColumn)->put_ColumnData(columnData);
		}
		if(stateImageIndex > 0 && properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
			// to set the state image index, we have to work directly on the header, so do it _after_ insertion
			(*ppAddedColumn)->put_StateImageIndex(stateImageIndex);
		}
		if(columnWidth < 0) {
			// the column can't be auto-sized on creation, so do it _after_ insertion
			(*ppAddedColumn)->put_Width(columnWidth);
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewColumns::ClearAllFilters(void)
{
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(::IsWindow(hWndHeader));

	if(SendMessage(hWndHeader, HDM_CLEARFILTER, static_cast<WPARAM>(-1), 0)) {
		return S_OK;
	}
	return E_FAIL;
}
STDMETHODIMP ListViewColumns::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));
	*pValue = static_cast<LONG>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
	if(*pValue == -1) {
		return E_FAIL;
	}

	return S_OK;
}

STDMETHODIMP ListViewColumns::Remove(LONG columnIdentifier, ColumnIdentifierTypeConstants columnIdentifierType/* = citIndex*/)
{
	// get the column's index
	int columnIndex = -1;
	switch(columnIdentifierType) {
		case citIndex:
			if(columnIdentifier >= 0) {
				HWND hWndHeader = properties.GetHeaderHWnd();
				ATLASSERT(IsWindow(hWndHeader));

				int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
				if(columnCount == -1) {
					return E_FAIL;
				}
				if(columnIdentifier < columnCount) {
					columnIndex = columnIdentifier;
				}
			}
			break;
		case citPosition:
			if(columnIdentifier >= 0) {
				HWND hWndHeader = properties.GetHeaderHWnd();
				ATLASSERT(IsWindow(hWndHeader));

				int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
				if(columnCount == -1) {
					return E_FAIL;
				}
				if(columnIdentifier < columnCount) {
					columnIndex = static_cast<int>(SendMessage(hWndHeader, HDM_ORDERTOINDEX, columnIdentifier, 0));
				}
			}
			break;
		case citID:
			if(properties.pOwnerExLvw) {
				columnIndex = properties.pOwnerExLvw->IDToColumnIndex(columnIdentifier);
			}
			break;
	}

	if(columnIndex != -1) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		// TODO: deleting column 0 isn't possible with comctl32.dll prior version 5.0
		if(SendMessage(hWndLvw, LVM_DELETECOLUMN, columnIndex, 0)) {
			return S_OK;
		}
	} else {
		// column not found
		return DISP_E_BADINDEX;
	}

	return E_FAIL;
}

STDMETHODIMP ListViewColumns::RemoveAll(void)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));
	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
	if(columnCount == -1) {
		return E_FAIL;
	}

	for(int i = columnCount - 1; i >= 0; --i) {
		// TODO: deleting column 0 isn't possible with comctl32.dll prior version 5.0
		SendMessage(hWndLvw, LVM_DELETECOLUMN, i, 0);
	}

	return S_OK;
}