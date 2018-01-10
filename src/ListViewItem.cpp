// ListViewItem.cpp: A wrapper for existing listview items.

#include "stdafx.h"
#include "ListViewItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListViewItem::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewItem::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HRESULT ListViewItem::Properties::WindowQueryInterface(REFIID requiredInterface, LPUNKNOWN* ppObject)
{
	ATLASSERT_POINTER(ppObject, LPUNKNOWN);
	if(!ppObject) {
		return E_POINTER;
	}

	*ppObject = NULL;

	HWND hWndLvw = GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));
	if(SendMessage(hWndLvw, LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&requiredInterface), reinterpret_cast<LPARAM>(ppObject))) {
		return S_OK;
	}
	return E_NOTIMPL;
}


HRESULT ListViewItem::SaveState(LVITEMINDEX& itemIndex, LPLVITEM pTarget, HWND hWndLvw/* = NULL*/)
{
	if(!hWndLvw) {
		hWndLvw = properties.GetExLvwHWnd();
	}
	ATLASSERT(IsWindow(hWndLvw));
	if(itemIndex.iItem < 0 || itemIndex.iItem >= static_cast<int>(SendMessage(hWndLvw, LVM_GETITEMCOUNT, 0, 0))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, LVITEM);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(LVITEM));
	pTarget->iItem = itemIndex.iItem;
	pTarget->cchTextMax = MAX_ITEMTEXTLENGTH;
	pTarget->pszText = reinterpret_cast<LPTSTR>(malloc((pTarget->cchTextMax + 1) * sizeof(TCHAR)));
	pTarget->puColumns = new UINT[200];
	pTarget->mask = LVIF_COLUMNS | LVIF_GROUPID | LVIF_IMAGE | LVIF_INDENT | LVIF_NORECOMPUTE | LVIF_PARAM | LVIF_STATE | LVIF_TEXT;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		pTarget->cColumns = 200;
		pTarget->piColFmt = new int[pTarget->cColumns];
		pTarget->mask |= LVIF_COLFMT;
		pTarget->iGroup = itemIndex.iGroup;
	}

	return static_cast<HRESULT>(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(pTarget)));
}


void ListViewItem::Attach(LVITEMINDEX& itemIndex)
{
	properties.itemIndex = itemIndex;
}

void ListViewItem::Detach(void)
{
	properties.itemIndex.iItem = -2;
	properties.itemIndex.iGroup = 0;
}

HRESULT ListViewItem::LoadState(LPLVITEM pSource)
{
	ATLASSERT_POINTER(pSource, LVITEM);
	if(!pSource) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(pSource->mask & LVIF_STATE) {
		LVITEM item = {0};
		if(pSource->state & LVIS_DROPHILITED) {
			item.state = LVIS_DROPHILITED;
		}
		item.stateMask = LVIS_DROPHILITED;
		if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
			SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item));
		} else {
			SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item));
		}
	}

	// now apply the settings
	VARIANT_BOOL b = VARIANT_FALSE;
	if(pSource->mask & LVIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->state & LVIS_ACTIVATING);
		put_Activating(b);
	}
	b = VARIANT_FALSE;
	if(pSource->mask & LVIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->state & LVIS_CUT);
		put_Ghosted(b);
	}
	b = VARIANT_FALSE;
	if(pSource->mask & LVIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->state & LVIS_GLOW);
		put_Glowing(b);
	}
	b = VARIANT_FALSE;
	if(pSource->mask & LVIF_STATE) {
		b = BOOL2VARIANTBOOL(pSource->state & LVIS_SELECTED);
		put_Selected(b);
	}

	BSTR text = NULL;
	if(pSource->mask & LVIF_TEXT) {
		if(pSource->pszText != LPSTR_TEXTCALLBACK) {
			text = _bstr_t(pSource->pszText).Detach();
		}
		put_Text(text);
	}
	SysFreeString(text);

	VARIANT v;
	VariantInit(&v);
	if((pSource->mask & LVIF_COLUMNS) == LVIF_COLUMNS || (pSource->mask & LVIF_COLFMT) == LVIF_COLFMT) {
		if(pSource->cColumns == I_COLUMNSCALLBACK) {
			// return 'Empty'
			v.vt = VT_EMPTY;
		} else if(pSource->cColumns > 0) {
			// return an array
			CComPtr<IRecordInfo> pRecordInfo = NULL;
			CLSID clsidTILEVIEWSUBITEM = {0};
			#ifdef UNICODE
				CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
			#else
				CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
			#endif
			v.parray = SafeArrayCreateVectorEx(VT_RECORD, 0, pSource->cColumns, pRecordInfo);
			v.vt = VT_ARRAY | VT_RECORD;

			TILEVIEWSUBITEM element = {0};
			for(LONG i = 0; i < static_cast<LONG>(pSource->cColumns); ++i) {
				element.ColumnIndex = pSource->puColumns[i];
				element.BeginNewColumn = BOOL2VARIANTBOOL(pSource->piColFmt[i] & LVCFMT_LINE_BREAK);
				element.FillRemainder = BOOL2VARIANTBOOL(pSource->piColFmt[i] & LVCFMT_FILL);
				element.WrapToMultipleLines = BOOL2VARIANTBOOL(pSource->piColFmt[i] & LVCFMT_WRAP);
				element.NoTitle = BOOL2VARIANTBOOL(pSource->piColFmt[i] & LVCFMT_NO_TITLE);
				ATLVERIFY(SUCCEEDED(SafeArrayPutElement(v.parray, &i, &element)));
			}
		} else {
			// return an empty array
			v.parray = NULL;
			v.vt = VT_ARRAY | VT_RECORD;
		}
		put_TileViewColumns(v);
	}
	VariantClear(&v);

	if(pSource->mask & LVIF_GROUPID) {
		CComPtr<IListViewGroup> pLvwGroup = ClassFactory::InitGroup(pSource->iGroupId, properties.pOwnerExLvw);
		putref_Group(pLvwGroup);
	}

	LONG l = 0;
	if(pSource->mask & LVIF_IMAGE) {
		l = pSource->iImage;
		put_IconIndex(l);
	}
	l = 0;
	if(pSource->mask & LVIF_INDENT) {
		l = pSource->iIndent;
		put_Indent(l);
	}
	l = 0;
	if(pSource->mask & LVIF_PARAM) {
		l = static_cast<LONG>(pSource->lParam);
		put_ItemData(l);
	}
	l = 0;
	if(pSource->mask & LVIF_STATE) {
		l = ((pSource->state & LVIS_OVERLAYMASK) >> 8);
		put_OverlayIndex(l);
	}
	l = 1;
	if(pSource->mask & LVIF_STATE) {
		l = ((pSource->state & LVIS_STATEIMAGEMASK) >> 12);
		put_StateImageIndex(l);
	}

	return S_OK;
}

HRESULT ListViewItem::LoadState(VirtualListViewItem* pSource)
{
	ATLASSUME(pSource);
	if(!pSource) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	VARIANT_BOOL b = VARIANT_FALSE;
	pSource->get_DropHilited(&b);
	LVITEM item = {0};
	if(b != VARIANT_FALSE) {
		item.state = LVIS_DROPHILITED;
	}
	item.stateMask = LVIS_DROPHILITED;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item));
	} else {
		SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item));
	}

	// now apply the settings
	b = VARIANT_FALSE;
	pSource->get_Activating(&b);
	put_Activating(b);
	b = VARIANT_FALSE;
	pSource->get_Ghosted(&b);
	put_Ghosted(b);
	b = VARIANT_FALSE;
	pSource->get_Glowing(&b);
	put_Glowing(b);
	b = VARIANT_FALSE;
	pSource->get_Selected(&b);
	put_Selected(b);

	BSTR text = NULL;
	pSource->get_Text(&text);
	put_Text(text);
	SysFreeString(text);

	VARIANT v;
	VariantInit(&v);
	pSource->get_TileViewColumns(&v);
	put_TileViewColumns(v);
	VariantClear(&v);

	CComPtr<IListViewGroup> pLvwGroup = NULL;
	pSource->get_Group(&pLvwGroup);
	putref_Group(pLvwGroup);

	LONG l = 0;
	pSource->get_IconIndex(&l);
	put_IconIndex(l);
	l = 0;
	pSource->get_Indent(&l);
	put_Indent(l);
	l = 0;
	pSource->get_ItemData(&l);
	put_ItemData(l);
	l = 0;
	pSource->get_OverlayIndex(&l);
	put_OverlayIndex(l);
	l = 1;
	pSource->get_StateImageIndex(&l);
	put_StateImageIndex(l);

	return S_OK;
}

HRESULT ListViewItem::SaveState(LPLVITEM pTarget, HWND hWndLvw/* = NULL*/)
{
	return SaveState(properties.itemIndex, pTarget, hWndLvw);
}

HRESULT ListViewItem::SaveState(VirtualListViewItem* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerExLvw);
	LVITEM item = {0};
	HRESULT hr = SaveState(&item);
	pTarget->LoadState(&item);
	SECUREFREE(item.pszText);

	return hr;
}


void ListViewItem::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewItem::get_Activating(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		*pValue = VARIANT_FALSE;
		return S_OK;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_ACTIVATING, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_ACTIVATING, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_ACTIVATING));
		hr = S_OK;
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_ACTIVATING);
	return hr;
}

STDMETHODIMP ListViewItem::put_Activating(VARIANT_BOOL newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	if(newValue != VARIANT_FALSE) {
		item.state = LVIS_ACTIVATING;
	}
	item.stateMask = LVIS_ACTIVATING;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		if(SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_Anchor(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		LVITEMINDEX selectionMark;
		if(SUCCEEDED(pListView7->GetSelectionMark(&selectionMark))) {
			*pValue = BOOL2VARIANTBOOL(properties.itemIndex.iItem == selectionMark.iItem && properties.itemIndex.iGroup == selectionMark.iGroup);
			hr = S_OK;
		}
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		LVITEMINDEX selectionMark;
		if(SUCCEEDED(pListViewVista->GetSelectionMark(&selectionMark))) {
			*pValue = BOOL2VARIANTBOOL(properties.itemIndex.iItem == selectionMark.iItem && properties.itemIndex.iGroup == selectionMark.iGroup);
			hr = S_OK;
		}
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		*pValue = BOOL2VARIANTBOOL(properties.itemIndex.iItem == static_cast<int>(SendMessage(hWndLvw, LVM_GETSELECTIONMARK, 0, 0)));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewItem::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_FOCUSED, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_FOCUSED, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_FOCUSED));
		hr = S_OK;
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_FOCUSED);
	return hr;
}

STDMETHODIMP ListViewItem::get_DropHilited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_DROPHILITED, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_DROPHILITED, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_DROPHILITED));
		hr = S_OK;
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_DROPHILITED);
	return hr;
}

STDMETHODIMP ListViewItem::get_Ghosted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		*pValue = VARIANT_FALSE;
		return S_OK;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_CUT, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_CUT, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_CUT));
		hr = S_OK;
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_CUT);
	return hr;
}

STDMETHODIMP ListViewItem::put_Ghosted(VARIANT_BOOL newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	if(newValue != VARIANT_FALSE) {
		item.state = LVIS_CUT;
	}
	item.stateMask = LVIS_CUT;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
		if(SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_Glowing(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		*pValue = VARIANT_FALSE;
		return S_OK;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_GLOW, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_GLOW, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_GLOW));
		hr = S_OK;
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_GLOW);
	return hr;
}

STDMETHODIMP ListViewItem::put_Glowing(VARIANT_BOOL newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	if(newValue != VARIANT_FALSE) {
		item.state = LVIS_GLOW;
	}
	item.stateMask = LVIS_GLOW;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		if(SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_Group(IListViewGroup** ppGroup)
{
	ATLASSERT_POINTER(ppGroup, IListViewGroup*);
	if(!ppGroup) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;

	if(RunTimeHelper::IsCommCtrl6()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.iItem = properties.itemIndex.iItem;
		item.iGroup = properties.itemIndex.iGroup;
		item.mask = LVIF_GROUPID;
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			ClassFactory::InitGroup(item.iGroupId, properties.pOwnerExLvw, IID_IListViewGroup, reinterpret_cast<LPUNKNOWN*>(ppGroup));
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewItem::putref_Group(IListViewGroup* pNewGroup)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	if(RunTimeHelper::IsCommCtrl6()) {
		int newGroupID = -2;
		if(pNewGroup) {
			LONG l = -2;
			pNewGroup->get_ID(&l);
			newGroupID = l;
		}

		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.iItem = properties.itemIndex.iItem;
		item.iGroupId = newGroupID;
		item.mask = LVIF_GROUPID;
		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_GroupIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex.iGroup;
	return S_OK;
}

STDMETHODIMP ListViewItem::get_Hot(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		LVITEMINDEX hotItem;
		if(SUCCEEDED(pListView7->GetHotItem(&hotItem))) {
			*pValue = BOOL2VARIANTBOOL(properties.itemIndex.iItem == hotItem.iItem && properties.itemIndex.iGroup == hotItem.iGroup);
			hr = S_OK;
		}
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		LVITEMINDEX hotItem;
		if(SUCCEEDED(pListViewVista->GetHotItem(&hotItem))) {
			*pValue = BOOL2VARIANTBOOL(properties.itemIndex.iItem == hotItem.iItem && properties.itemIndex.iGroup == hotItem.iGroup);
			hr = S_OK;
		}
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		*pValue = BOOL2VARIANTBOOL(properties.itemIndex.iItem == static_cast<int>(SendMessage(hWndLvw, LVM_GETHOTITEM, 0, 0)));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.itemIndex.iItem;
	item.iGroup = properties.itemIndex.iGroup;
	item.mask = LVIF_IMAGE;
	if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::put_IconIndex(LONG newValue)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	//if(properties.pOwnerExLvw->IsComctl32Version581OrNewer()) {
		if(newValue < -2) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	/*} else {
		if(newValue < -1) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	}*/

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.itemIndex.iItem;
	item.iGroup = properties.itemIndex.iGroup;
	item.iImage = newValue;
	item.mask = LVIF_IMAGE;
	if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	if(properties.pOwnerExLvw) {
		*pValue = properties.pOwnerExLvw->ItemIndexToID(properties.itemIndex.iItem);
		if(*pValue != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_Indent(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.itemIndex.iItem;
	item.iGroup = properties.itemIndex.iGroup;
	item.mask = LVIF_INDENT;
	if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iIndent;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::put_Indent(OLE_XSIZE_PIXELS newValue)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.itemIndex.iItem;
	item.iGroup = properties.itemIndex.iGroup;
	item.iIndent = newValue;
	item.mask = LVIF_INDENT;
	if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex.iItem;
	return S_OK;
}

STDMETHODIMP ListViewItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.itemIndex.iItem;
	item.iGroup = properties.itemIndex.iGroup;
	item.mask = LVIF_PARAM;
	if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = static_cast<LONG>(item.lParam);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::put_ItemData(LONG newValue)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.itemIndex.iItem;
	item.iGroup = properties.itemIndex.iGroup;
	item.lParam = newValue;
	item.mask = LVIF_PARAM;
	if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_OVERLAYMASK, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_OVERLAYMASK, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_OVERLAYMASK));
		hr = S_OK;
	}
	*pValue = (state & LVIS_OVERLAYMASK) >> 8;
	return hr;
}

STDMETHODIMP ListViewItem::put_OverlayIndex(LONG newValue)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.state = INDEXTOOVERLAYMASK(newValue);
	item.stateMask = LVIS_OVERLAYMASK;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		if(SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item))) {
			// seems to be necessary
			SendMessage(hWndLvw, LVM_REDRAWITEMS, properties.itemIndex.iItem, properties.itemIndex.iItem);
			return S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			// seems to be necessary
			SendMessage(hWndLvw, LVM_REDRAWITEMS, properties.itemIndex.iItem, properties.itemIndex.iItem);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_Selected(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		*pValue = VARIANT_FALSE;
		return S_OK;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_SELECTED, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_SELECTED, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_SELECTED));
		hr = S_OK;
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_SELECTED);
	return hr;
}

STDMETHODIMP ListViewItem::put_Selected(VARIANT_BOOL newValue)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	if(newValue != VARIANT_FALSE) {
		item.state = LVIS_SELECTED;
	}
	item.stateMask = LVIS_SELECTED;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		if(SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	STDMETHODIMP ListViewItem::get_ShellListViewItemObject(IDispatch** ppItem)
	{
		ATLASSERT_POINTER(ppItem, LPDISPATCH);
		if(!ppItem) {
			return E_POINTER;
		}

		if(properties.itemIndex.iItem < 0) {
			return E_FAIL;
		}

		if(properties.pOwnerExLvw) {
			if(properties.pOwnerExLvw->shellBrowserInterface.pInternalMessageListener) {
				LONG itemID = properties.pOwnerExLvw->ItemIndexToID(properties.itemIndex.iItem);
				if(itemID != -1) {
					properties.pOwnerExLvw->shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_GETSHLISTVIEWITEM, itemID, reinterpret_cast<LPARAM>(ppItem));
					return S_OK;
				}
			}
		}
		return E_FAIL;
	}
#endif

STDMETHODIMP ListViewItem::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		*pValue = -1;
		return S_OK;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.itemIndex.iItem, 0, LVIS_STATEIMAGEMASK, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.itemIndex.iItem, 0, LVIS_STATEIMAGEMASK, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		state = static_cast<ULONG>(SendMessage(hWndLvw, LVM_GETITEMSTATE, properties.itemIndex.iItem, LVIS_STATEIMAGEMASK));
		hr = S_OK;
	}
	*pValue = (state & LVIS_STATEIMAGEMASK) >> 12;
	return hr;
}

STDMETHODIMP ListViewItem::put_StateImageIndex(LONG newValue)
{
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.state = INDEXTOSTATEIMAGEMASK(newValue);
	item.stateMask = LVIS_STATEIMAGEMASK;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		if(SendMessage(hWndLvw, LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_SETITEMSTATE, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::get_SubItems(IListViewSubItems** ppSubItems)
 {
	ATLASSERT_POINTER(ppSubItems, IListViewSubItems*);
	if(!ppSubItems) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	ClassFactory::InitListSubItems(properties.itemIndex, properties.pOwnerExLvw, IID_IListViewSubItems, reinterpret_cast<LPUNKNOWN*>(ppSubItems));
	return S_OK;
}

STDMETHODIMP ListViewItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	LPWSTR pBuffer = reinterpret_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (MAX_ITEMTEXTLENGTH + 1) * sizeof(WCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		//if(SUCCEEDED(pListView7->GetItemText(properties.itemIndex.iItem, 0, pBuffer, MAX_ITEMTEXTLENGTH))) {
		// returns E_FAIL for empty sub-items?!
		pListView7->GetItemText(properties.itemIndex.iItem, 0, pBuffer, MAX_ITEMTEXTLENGTH);
			*pValue = _bstr_t(pBuffer).Detach();
			hr = S_OK;
		//}
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		//if(SUCCEEDED(pListViewVista->GetItemText(properties.itemIndex.iItem, 0, pBuffer, MAX_ITEMTEXTLENGTH))) {
		// returns E_FAIL for empty sub-items?!
		pListViewVista->GetItemText(properties.itemIndex.iItem, 0, pBuffer, MAX_ITEMTEXTLENGTH);
			*pValue = _bstr_t(pBuffer).Detach();
			hr = S_OK;
		//}
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.cchTextMax = MAX_ITEMTEXTLENGTH;
		item.pszText = reinterpret_cast<LPTSTR>(pBuffer);
		SendMessage(hWndLvw, LVM_GETITEMTEXT, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item));

		*pValue = _bstr_t(item.pszText).Detach();
		hr = S_OK;
	}
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return hr;
}

STDMETHODIMP ListViewItem::put_Text(BSTR newValue)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		if(newValue) {
			hr = pListView7->SetItemText(properties.itemIndex.iItem, 0, OLE2W(newValue));
		} else {
			hr = pListView7->SetItemText(properties.itemIndex.iItem, 0, LPSTR_TEXTCALLBACKW);
		}
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		if(newValue) {
			hr = pListViewVista->SetItemText(properties.itemIndex.iItem, 0, OLE2W(newValue));
		} else {
			hr = pListViewVista->SetItemText(properties.itemIndex.iItem, 0, LPSTR_TEXTCALLBACKW);
		}
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		COLE2T converter(newValue);
		if(newValue) {
			item.pszText = converter;
		} else {
			item.pszText = LPSTR_TEXTCALLBACK;
		}
		if(SendMessage(hWndLvw, LVM_SETITEMTEXT, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewItem::get_TileViewColumns(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	if(RunTimeHelper::IsCommCtrl6()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.iItem = properties.itemIndex.iItem;
		item.iGroup = properties.itemIndex.iGroup;
		item.mask = LVIF_COLUMNS | LVIF_COLFMT;
		// 200 should be large enough ;)
		item.cColumns = 200;
		item.puColumns = new UINT[item.cColumns];
		item.piColFmt = new int[item.cColumns];
		ZeroMemory(item.piColFmt, item.cColumns * sizeof(int));
		if(!properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
			// yes, comctl32.dll 6.0 requires this field to be 0
			item.cColumns = 0;
		}
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			if(item.cColumns > 0) {
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
				pValue->parray = SafeArrayCreateVectorEx(VT_RECORD, 0, item.cColumns, pRecordInfo);
				pValue->vt = VT_ARRAY | VT_RECORD;

				TILEVIEWSUBITEM element = {0};
				for(LONG i = 0; i < static_cast<LONG>(item.cColumns); ++i) {
					element.ColumnIndex = item.puColumns[i];
					element.BeginNewColumn = BOOL2VARIANTBOOL(item.piColFmt[i] & LVCFMT_LINE_BREAK);
					element.FillRemainder = BOOL2VARIANTBOOL(item.piColFmt[i] & LVCFMT_FILL);
					element.WrapToMultipleLines = BOOL2VARIANTBOOL(item.piColFmt[i] & LVCFMT_WRAP);
					element.NoTitle = BOOL2VARIANTBOOL(item.piColFmt[i] & LVCFMT_NO_TITLE);
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

		delete[] item.puColumns;
		delete[] item.piColFmt;
	}
	return hr;
}

STDMETHODIMP ListViewItem::put_TileViewColumns(VARIANT newValue)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	if(RunTimeHelper::IsCommCtrl6()) {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.iItem = properties.itemIndex.iItem;
		item.iGroup = properties.itemIndex.iGroup;
		item.mask = LVIF_COLUMNS | LVIF_COLFMT;

		CLSID clsidTILEVIEWSUBITEM;
		#ifdef UNICODE
			CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
		#else
			CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
		#endif
		if(newValue.vt == VT_EMPTY) {
			item.cColumns = I_COLUMNSCALLBACK;
		} else if(newValue.vt & VT_ARRAY) {
			// an array
			if(newValue.parray) {
				LONG l = 0;
				SafeArrayGetLBound(newValue.parray, 1, &l);
				LONG u = 0;
				SafeArrayGetUBound(newValue.parray, 1, &u);
				if(u < l) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				VARTYPE vt = 0;
				SafeArrayGetVartype(newValue.parray, &vt);
				if(vt == VT_VARIANT) {
					for(LONG i = l; i <= u; ++i) {
						VARIANT buffer;
						SafeArrayGetElement(newValue.parray, &i, &buffer);
						if(buffer.vt != VT_RECORD) {
							// invalid arg - raise VB runtime error 380
							VariantClear(&buffer);
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}
						if(!buffer.pRecInfo) {
							// invalid arg - raise VB runtime error 380
							VariantClear(&buffer);
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}
						GUID recordGuid = {0};
						ATLVERIFY(SUCCEEDED(buffer.pRecInfo->GetGuid(&recordGuid)));
						ULONG elementSize = 0;
						ATLVERIFY(SUCCEEDED(buffer.pRecInfo->GetSize(&elementSize)));
						if(recordGuid != clsidTILEVIEWSUBITEM || elementSize > sizeof(TILEVIEWSUBITEM)) {
							// invalid arg - raise VB runtime error 380
							VariantClear(&buffer);
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}
						VariantClear(&buffer);
					}
				} else if(vt == VT_RECORD) {
					CComPtr<IRecordInfo> pRecordInfo = NULL;
					ATLVERIFY(SUCCEEDED(SafeArrayGetRecordInfo(newValue.parray, &pRecordInfo)));
					if(!pRecordInfo) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					GUID recordGuid = {0};
					ATLVERIFY(SUCCEEDED(pRecordInfo->GetGuid(&recordGuid)));
					ULONG elementSize = 0;
					ATLVERIFY(SUCCEEDED(pRecordInfo->GetSize(&elementSize)));
					if(recordGuid != clsidTILEVIEWSUBITEM || elementSize > sizeof(TILEVIEWSUBITEM)) {
						// invalid value - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
				} else {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}

				item.cColumns = u - l + 1;
				item.puColumns = new UINT[item.cColumns];
				item.piColFmt = new int[item.cColumns];

				TILEVIEWSUBITEM element = {0};
				for(LONG i = l; i <= u; ++i) {
					if(vt == VT_RECORD) {
						ATLVERIFY(SUCCEEDED(SafeArrayGetElement(newValue.parray, &i, &element)));
					} else {
						VARIANT v;
						ATLVERIFY(SUCCEEDED(SafeArrayGetElement(newValue.parray, &i, &v)));
						element = *reinterpret_cast<TILEVIEWSUBITEM*>(v.pvRecord);
						VariantClear(&v);
					}
					item.puColumns[i - l] = element.ColumnIndex;
					item.piColFmt[i - l] = 0;
					if(element.BeginNewColumn != VARIANT_FALSE) {
						item.piColFmt[i - l] |= LVCFMT_LINE_BREAK;
					}
					if(element.FillRemainder != VARIANT_FALSE) {
						item.piColFmt[i - l] |= LVCFMT_FILL;
					}
					if(element.WrapToMultipleLines != VARIANT_FALSE) {
						item.piColFmt[i - l] |= LVCFMT_WRAP;
					}
					if(element.NoTitle != VARIANT_FALSE) {
						item.piColFmt[i - l] |= LVCFMT_NO_TITLE;
					}
				}
			} else {
				// an empty array
				item.cColumns = 0;
			}
		} else if(newValue.vt == VT_RECORD) {
			// a single value
			if(!newValue.pRecInfo) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			GUID recordGuid = {0};
			ATLVERIFY(SUCCEEDED(newValue.pRecInfo->GetGuid(&recordGuid)));
			ULONG elementSize = 0;
			ATLVERIFY(SUCCEEDED(newValue.pRecInfo->GetSize(&elementSize)));
			if(recordGuid != clsidTILEVIEWSUBITEM || elementSize > sizeof(TILEVIEWSUBITEM)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			TILEVIEWSUBITEM* pData = reinterpret_cast<TILEVIEWSUBITEM*>(newValue.pvRecord);
			ATLASSERT_POINTER(pData, TILEVIEWSUBITEM);

			item.cColumns = 1;
			item.puColumns = new UINT[item.cColumns];
			item.piColFmt = new int[item.cColumns];

			item.puColumns[0] = pData->ColumnIndex;
			item.piColFmt[0] = 0;
			if(pData->BeginNewColumn != VARIANT_FALSE) {
				item.piColFmt[0] |= LVCFMT_LINE_BREAK;
			}
			if(pData->FillRemainder != VARIANT_FALSE) {
				item.piColFmt[0] |= LVCFMT_FILL;
			}
			if(pData->WrapToMultipleLines != VARIANT_FALSE) {
				item.piColFmt[0] |= LVCFMT_WRAP;
			}
			if(pData->NoTitle != VARIANT_FALSE) {
				item.piColFmt[0] |= LVCFMT_NO_TITLE;
			}
		}

		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
		if(item.puColumns) {
			delete[] item.puColumns;
			delete[] item.piColFmt;
		}
	}

	return hr;
}

STDMETHODIMP ListViewItem::get_WorkArea(IListViewWorkArea** ppWorkArea)
{
	ATLASSERT_POINTER(ppWorkArea, IListViewWorkArea*);
	if(!ppWorkArea) {
		return E_POINTER;
	}
	*ppWorkArea = NULL;

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	// get the working areas' rectangles
	int workAreas = 0;
	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(&workAreas));
	if(workAreas > 0) {
		LPRECT pWorkAreaRectangles = new RECT[workAreas];
		SendMessage(hWndLvw, LVM_GETWORKAREAS, workAreas, reinterpret_cast<LPARAM>(pWorkAreaRectangles));

		// now get the item's position
		POINT upperLeftPoint = {0};
		HRESULT hr = E_FAIL;
		CComPtr<IListView_WIN7> pListView7 = NULL;
		CComPtr<IListView_WINVISTA> pListViewVista = NULL;
		if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
			hr = pListView7->GetItemPosition(properties.itemIndex, &upperLeftPoint);
		} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
			hr = pListViewVista->GetItemPosition(properties.itemIndex, &upperLeftPoint);
		} else {
			if(SendMessage(hWndLvw, LVM_GETITEMPOSITION, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&upperLeftPoint))) {
				hr = S_OK;
			}
		}
		if(SUCCEEDED(hr)) {
			// search for the working area that we're in
			for(int i = 0; i < workAreas; ++i) {
				if(PtInRect(&pWorkAreaRectangles[i], upperLeftPoint)) {
					// found it
					ClassFactory::InitWorkArea(i, properties.pOwnerExLvw, IID_IListViewWorkArea, reinterpret_cast<LPUNKNOWN*>(ppWorkArea));
					delete[] pWorkAreaRectangles;
					return S_OK;
				}
			}
			// no specific working area seems to mean we're in the first one
			ClassFactory::InitWorkArea(0, properties.pOwnerExLvw, IID_IListViewWorkArea, reinterpret_cast<LPUNKNOWN*>(ppWorkArea));
		}
		delete[] pWorkAreaRectangles;
	}

	return S_OK;
}


STDMETHODIMP ListViewItem::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	ATLASSUME(properties.pOwnerExLvw);

	POINT upperLeftPoint = {0};
	*phImageList = HandleToLong(properties.pOwnerExLvw->CreateLegacyDragImage(properties.itemIndex, &upperLeftPoint, NULL));

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

STDMETHODIMP ListViewItem::EnsureVisible(VARIANT_BOOL mustBeEntirelyVisible/* = VARIANT_TRUE*/)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	// TODO: Don't we have to inverse mustBeEntirelyVisible?

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->EnsureItemVisible(properties.itemIndex, VARIANTBOOL2BOOL(mustBeEntirelyVisible));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->EnsureItemVisible(properties.itemIndex, VARIANTBOOL2BOOL(mustBeEntirelyVisible));
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		if(SendMessage(hWndLvw, LVM_ENSUREVISIBLE, properties.itemIndex.iItem, VARIANTBOOL2BOOL(mustBeEntirelyVisible))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewItem::FindNextItem(SearchModeConstants searchMode, VARIANT searchFor, SearchDirectionConstants searchDirection/* = sdNoneSpecific*/, VARIANT_BOOL wrapAtLastItem/* = VARIANT_TRUE*/, IListViewItem** ppFoundItem/* = NULL*/)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(FAILED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7)))) {
		properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista));
	}
	if(pListView7 || pListViewVista) {
		LVFINDINFOW findInfo = {0};
		VARIANT v;
		VariantInit(&v);
		switch(searchMode) {
			case smItemData:
				findInfo.flags = LVFI_PARAM;
				if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_UI4))) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				findInfo.lParam = v.ulVal;
				break;
			case smPartialText:
				findInfo.flags = LVFI_PARTIAL;
				// fall through
			case smText:
				findInfo.flags |= LVFI_STRING;
				if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_BSTR))) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				findInfo.psz = OLE2W(v.bstrVal);
				break;
			case smNearestPosition:
				findInfo.flags = LVFI_NEARESTXY;
				findInfo.vkDirection = static_cast<UINT>(searchDirection);
				if((searchFor.vt & VT_ARRAY) == VT_ARRAY && searchFor.parray) {
					LONG l = 0;
					SafeArrayGetLBound(searchFor.parray, 1, &l);
					LONG u = 0;
					SafeArrayGetUBound(searchFor.parray, 1, &u);
					if(u < l || u - l + 1 != 2) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}

					VARTYPE vt = 0;
					SafeArrayGetVartype(searchFor.parray, &vt);
					ULONG element = 0;
					for(LONG i = l; i <= u; ++i) {
						if(vt == VT_VARIANT) {
							SafeArrayGetElement(searchFor.parray, &i, &v);
							if(FAILED(VariantChangeType(&v, &v, 0, VT_UI4))) {
								// invalid arg - raise VB runtime error 380
								return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
							}
							element = v.ulVal;
						} else {
							SafeArrayGetElement(searchFor.parray, &i, &element);
						}
						if(i == l) {
							findInfo.pt.x = element;
						} else {
							findInfo.pt.y = element;
						}
					}
				} else {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
		}
		VariantClear(&v);
		if(wrapAtLastItem != VARIANT_FALSE) {
			findInfo.flags |= LVFI_WRAP;
		}

		LVITEMINDEX foundItem = {-1, 0};
		if(pListView7) {
			hr = pListView7->FindItem(properties.itemIndex, &findInfo, &foundItem);
		} else {
			hr = pListViewVista->FindItem(properties.itemIndex, &findInfo, &foundItem);
		}
		if(SUCCEEDED(hr)) {
			ClassFactory::InitListItem(foundItem, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
		} else {
			*ppFoundItem = NULL;
		}
		hr = S_OK;
	}
	if(FAILED(hr)) {
		LVFINDINFO findInfo = {0};
		VARIANT v;
		VariantInit(&v);
		switch(searchMode) {
			case smItemData:
				findInfo.flags = LVFI_PARAM;
				if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_UI4))) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				findInfo.lParam = v.ulVal;
				break;
			case smPartialText:
				findInfo.flags = LVFI_PARTIAL;
				// fall through
			case smText:
				findInfo.flags |= LVFI_STRING;
				if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_BSTR))) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case smNearestPosition:
				findInfo.flags = LVFI_NEARESTXY;
				findInfo.vkDirection = static_cast<UINT>(searchDirection);
				if((searchFor.vt & VT_ARRAY) == VT_ARRAY && searchFor.parray) {
					LONG l = 0;
					SafeArrayGetLBound(searchFor.parray, 1, &l);
					LONG u = 0;
					SafeArrayGetUBound(searchFor.parray, 1, &u);
					if(u < l || u - l + 1 != 2) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}

					VARTYPE vt = 0;
					SafeArrayGetVartype(searchFor.parray, &vt);
					ULONG element = 0;
					for(LONG i = l; i <= u; ++i) {
						if(vt == VT_VARIANT) {
							SafeArrayGetElement(searchFor.parray, &i, &v);
							if(FAILED(VariantChangeType(&v, &v, 0, VT_UI4))) {
								// invalid arg - raise VB runtime error 380
								return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
							}
							element = v.ulVal;
						} else {
							SafeArrayGetElement(searchFor.parray, &i, &element);
						}
						if(i == l) {
							findInfo.pt.x = element;
						} else {
							findInfo.pt.y = element;
						}
					}
				} else {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
		}
		COLE2T converter(v.bstrVal);
		if(findInfo.flags & LVFI_STRING) {
			findInfo.psz = converter;
		}
		if(wrapAtLastItem != VARIANT_FALSE) {
			findInfo.flags |= LVFI_WRAP;
		}

		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEMINDEX foundItem;
		foundItem.iItem = static_cast<int>(SendMessage(hWndLvw, LVM_FINDITEM, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&findInfo)));
		foundItem.iGroup = 0;
		ClassFactory::InitListItem(foundItem, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
		VariantClear(&v);
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewItem::GetPosition(OLE_XPOS_PIXELS* pX, OLE_YPOS_PIXELS* pY)
{
	ATLASSERT_POINTER(pX, OLE_XPOS_PIXELS);
	if(!pX) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(pY, OLE_XPOS_PIXELS);
	if(!pY) {
		return E_POINTER;
	}

	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	POINT upperLeftPoint = {0};
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemPosition(properties.itemIndex, &upperLeftPoint);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemPosition(properties.itemIndex, &upperLeftPoint);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		if(SendMessage(hWndLvw, LVM_GETITEMPOSITION, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&upperLeftPoint))) {
			hr = S_OK;
		}
	}

	if(SUCCEEDED(hr)) {
		if(pX) {
			*pX = upperLeftPoint.x;
		}
		if(pY) {
			*pY = upperLeftPoint.y;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewItem::GetRectangle(ItemRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	HRESULT hr = E_FAIL;
	RECT rc = {0};
	rc.left = rectangleType;
	if(properties.pOwnerExLvw->IsComctl32Version610OrNewer() && SendMessage(hWndLvw, LVM_ISGROUPVIEWENABLED, 0, 0) && (CWindow(hWndLvw).GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		if(SendMessage(hWndLvw, LVM_GETITEMINDEXRECT, reinterpret_cast<WPARAM>(&properties.itemIndex), reinterpret_cast<LPARAM>(&rc))) {
			hr = S_OK;
		}
	} else {
		if(SendMessage(hWndLvw, LVM_GETITEMRECT, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&rc))) {
			hr = S_OK;
		}
	}
	if(SUCCEEDED(hr)) {
		if(pXLeft) {
			*pXLeft = rc.left;
		}
		if(pYTop) {
			*pYTop = rc.top;
		}
		if(pXRight) {
			*pXRight = rc.right;
		}
		if(pYBottom) {
			*pYBottom = rc.bottom;
		}
	}
	return hr;
}

STDMETHODIMP ListViewItem::IsVisible(VARIANT_BOOL* pVisible)
{
	ATLASSERT_POINTER(pVisible, VARIANT_BOOL);
	if(!pVisible) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		BOOL visible;
		hr = pListView7->IsItemVisible(properties.itemIndex, &visible);
		*pVisible = BOOL2VARIANTBOOL(visible);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		BOOL visible;
		hr = pListViewVista->IsItemVisible(properties.itemIndex, &visible);
		*pVisible = BOOL2VARIANTBOOL(visible);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		*pVisible = BOOL2VARIANTBOOL(SendMessage(hWndLvw, LVM_ISITEMVISIBLE, properties.itemIndex.iItem, 0));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewItem::SetInfoTipText(BSTR text, VARIANT_BOOL* pSuccess)
{
	ATLASSERT_POINTER(pSuccess, VARIANT_BOOL);
	if(!pSuccess) {
		return E_POINTER;
	}
	if(!text) {
		return E_INVALIDARG;
	}

	LVSETINFOTIP infoTip = {0};
	infoTip.cbSize = sizeof(LVSETINFOTIP);
	infoTip.iItem = properties.itemIndex.iItem;
	infoTip.pszText = OLE2W(text);

	// Don't use IListView here, because the SettingItemInfoTipText event relies on LVM_SETINFOTIP.
	/*CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		*pSuccess = BOOL2VARIANTBOOL(SUCCEEDED(pListView7->SetInfoTip(&infoTip)));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		*pSuccess = BOOL2VARIANTBOOL(SUCCEEDED(pListViewVista->SetInfoTip(&infoTip)));
	} else {*/
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		*pSuccess = BOOL2VARIANTBOOL(SendMessage(hWndLvw, LVM_SETINFOTIP, 0, reinterpret_cast<LPARAM>(&infoTip)));
	//}
	return S_OK;
}

STDMETHODIMP ListViewItem::SetPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	POINT upperLeftPoint = {x, y};
	SendMessage(hWndLvw, LVM_SETITEMPOSITION32, properties.itemIndex.iItem, reinterpret_cast<LPARAM>(&upperLeftPoint));
	return S_OK;
}

STDMETHODIMP ListViewItem::StartLabelEditing(void)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		HWND hWndEdit;
		hr = pListView7->EditLabel(properties.itemIndex, NULL, &hWndEdit);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		HWND hWndEdit;
		hr = pListViewVista->EditLabel(properties.itemIndex, NULL, &hWndEdit);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		if(SendMessage(hWndLvw, LVM_EDITLABEL, properties.itemIndex.iItem, 0)) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewItem::Update(void)
{
	if(properties.itemIndex.iItem < 0) {
		return E_FAIL;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(SendMessage(hWndLvw, LVM_UPDATE, properties.itemIndex.iItem, 0)) {
		return S_OK;
	}
	return E_FAIL;
}