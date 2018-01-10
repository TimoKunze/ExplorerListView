// ListViewSubItem.cpp: A wrapper for existing listview sub-items.

#include "stdafx.h"
#include "ListViewSubItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewSubItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewSubItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListViewSubItem::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewSubItem::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HRESULT ListViewSubItem::Properties::WindowQueryInterface(REFIID requiredInterface, LPUNKNOWN* ppObject)
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


void ListViewSubItem::Attach(LVITEMINDEX& parentItemIndex, int subItemIndex)
{
	properties.parentItemIndex = parentItemIndex;
	properties.subItemIndex = subItemIndex;
}

void ListViewSubItem::Detach(void)
{
	properties.parentItemIndex.iItem = -1;
	properties.parentItemIndex.iGroup = 0;
	properties.subItemIndex = -1;
}

void ListViewSubItem::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewSubItem::get_Activating(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_ACTIVATING, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_ACTIVATING, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_ACTIVATING;
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			state = item.state;
			hr = S_OK;
		}
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_ACTIVATING);
	return hr;
}

STDMETHODIMP ListViewSubItem::put_Activating(VARIANT_BOOL newValue)
{
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_ACTIVATING, (newValue != VARIANT_FALSE ? LVIS_ACTIVATING : 0));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_ACTIVATING, (newValue != VARIANT_FALSE ? LVIS_ACTIVATING : 0));
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_ACTIVATING;
		if(newValue != VARIANT_FALSE) {
			item.state = LVIS_ACTIVATING;
		}
		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewSubItem::get_Ghosted(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_CUT, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_CUT, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_CUT;
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			state = item.state;
			hr = S_OK;
		}
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_CUT);
	return hr;
}

STDMETHODIMP ListViewSubItem::put_Ghosted(VARIANT_BOOL newValue)
{
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_CUT, (newValue != VARIANT_FALSE ? LVIS_CUT : 0));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_CUT, (newValue != VARIANT_FALSE ? LVIS_CUT : 0));
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_CUT;
		if(newValue != VARIANT_FALSE) {
			item.state = LVIS_CUT;
		}
		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewSubItem::get_Glowing(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_GLOW, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_GLOW, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_GLOW;
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			state = item.state;
			hr = S_OK;
		}
	}
	*pValue = BOOL2VARIANTBOOL(state & LVIS_GLOW);
	return hr;
}

STDMETHODIMP ListViewSubItem::put_Glowing(VARIANT_BOOL newValue)
{
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_GLOW, (newValue != VARIANT_FALSE ? LVIS_GLOW : 0));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_GLOW, (newValue != VARIANT_FALSE ? LVIS_GLOW : 0));
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_GLOW;
		if(newValue != VARIANT_FALSE) {
			item.state = LVIS_GLOW;
		}
		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewSubItem::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVITEM item = {0};
	item.iItem = properties.parentItemIndex.iItem;
	item.iGroup = properties.parentItemIndex.iGroup;
	item.iSubItem = properties.subItemIndex;
	item.mask = LVIF_IMAGE;
	if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		*pValue = item.iImage;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewSubItem::put_IconIndex(LONG newValue)
{
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
	item.iItem = properties.parentItemIndex.iItem;
	item.iGroup = properties.parentItemIndex.iGroup;
	item.iSubItem = properties.subItemIndex;
	item.iImage = newValue;
	item.mask = LVIF_IMAGE;
	if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewSubItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.subItemIndex;
	return S_OK;
}

STDMETHODIMP ListViewSubItem::get_OverlayIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_OVERLAYMASK, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_OVERLAYMASK, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_OVERLAYMASK;
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			state = item.state;
			hr = S_OK;
		}
	}
	*pValue = (state & LVIS_OVERLAYMASK) >> 8;
	return hr;
}

STDMETHODIMP ListViewSubItem::put_OverlayIndex(LONG newValue)
{
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_OVERLAYMASK, INDEXTOOVERLAYMASK(newValue));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_OVERLAYMASK, INDEXTOOVERLAYMASK(newValue));
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_OVERLAYMASK;
		item.state = INDEXTOOVERLAYMASK(newValue);
		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewSubItem::get_ParentItem(IListViewItem** ppParentItem)
{
	ATLASSERT_POINTER(ppParentItem, IListViewItem*);
	if(!ppParentItem) {
		return E_POINTER;
	}

	ClassFactory::InitListItem(properties.parentItemIndex, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppParentItem));
	return S_OK;
}

STDMETHODIMP ListViewSubItem::get_StateImageIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	ULONG state = 0;
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_STATEIMAGEMASK, &state);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_STATEIMAGEMASK, &state);
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_STATEIMAGEMASK;
		if(SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			state = item.state;
			hr = S_OK;
		}
	}
	*pValue = (state & LVIS_STATEIMAGEMASK) >> 12;
	return hr;
}

STDMETHODIMP ListViewSubItem::put_StateImageIndex(LONG newValue)
{
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_STATEIMAGEMASK, INDEXTOSTATEIMAGEMASK(newValue));
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->SetItemState(properties.parentItemIndex.iItem, properties.subItemIndex, LVIS_STATEIMAGEMASK, INDEXTOSTATEIMAGEMASK(newValue));
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));
		LVITEM item = {0};
		item.iItem = properties.parentItemIndex.iItem;
		item.iSubItem = properties.subItemIndex;
		item.mask = LVIF_STATE;
		item.stateMask = LVIS_STATEIMAGEMASK;
		item.state = INDEXTOSTATEIMAGEMASK(newValue);
		if(SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ListViewSubItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	LPWSTR pBuffer = reinterpret_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (MAX_ITEMTEXTLENGTH + 1) * sizeof(WCHAR)));
	if(!pBuffer) {
		return E_OUTOFMEMORY;
	}

	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		//if(SUCCEEDED(pListView7->GetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, pBuffer, MAX_ITEMTEXTLENGTH))) {
		// returns E_FAIL for empty sub-items?!
		pListView7->GetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, pBuffer, MAX_ITEMTEXTLENGTH);
			*pValue = _bstr_t(pBuffer).Detach();
			hr = S_OK;
		//}
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		//if(SUCCEEDED(pListViewVista->GetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, pBuffer, MAX_ITEMTEXTLENGTH))) {
		// returns E_FAIL for empty sub-items?!
		pListViewVista->GetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, pBuffer, MAX_ITEMTEXTLENGTH);
			*pValue = _bstr_t(pBuffer).Detach();
			hr = S_OK;
		//}
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.iSubItem = properties.subItemIndex;
		item.cchTextMax = MAX_ITEMTEXTLENGTH;
		item.pszText = reinterpret_cast<LPTSTR>(pBuffer);
		SendMessage(hWndLvw, LVM_GETITEMTEXT, properties.parentItemIndex.iItem, reinterpret_cast<LPARAM>(&item));

		*pValue = _bstr_t(item.pszText).Detach();
		hr = S_OK;
	}
	HeapFree(GetProcessHeap(), 0, pBuffer);
	return hr;
}

STDMETHODIMP ListViewSubItem::put_Text(BSTR newValue)
{
	HRESULT hr = E_FAIL;
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		if(newValue) {
			hr = pListView7->SetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, OLE2W(newValue));
		} else {
			hr = pListView7->SetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, LPSTR_TEXTCALLBACKW);
		}
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		if(newValue) {
			hr = pListViewVista->SetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, OLE2W(newValue));
		} else {
			hr = pListViewVista->SetItemText(properties.parentItemIndex.iItem, properties.subItemIndex, LPSTR_TEXTCALLBACKW);
		}
	} else {
		HWND hWndLvw = properties.GetExLvwHWnd();
		ATLASSERT(IsWindow(hWndLvw));

		LVITEM item = {0};
		item.iSubItem = properties.subItemIndex;
		COLE2T converter(newValue);
		if(newValue) {
			item.pszText = converter;
		} else {
			item.pszText = LPSTR_TEXTCALLBACK;
		}
		if(SendMessage(hWndLvw, LVM_SETITEMTEXT, properties.parentItemIndex.iItem, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}
	return hr;
}


STDMETHODIMP ListViewSubItem::GetRectangle(ItemRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	HRESULT hr = E_FAIL;
	RECT rc = {0};
	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WIN7, reinterpret_cast<LPUNKNOWN*>(&pListView7))) && pListView7) {
		hr = pListView7->GetSubItemRect(properties.parentItemIndex, properties.subItemIndex, rectangleType, &rc);
	} else if(SUCCEEDED(properties.WindowQueryInterface(IID_IListView_WINVISTA, reinterpret_cast<LPUNKNOWN*>(&pListViewVista))) && pListViewVista) {
		hr = pListViewVista->GetSubItemRect(properties.parentItemIndex, properties.subItemIndex, rectangleType, &rc);
	} else {
		rc.left = rectangleType;
		rc.top = properties.subItemIndex;
		if(SendMessage(hWndLvw, LVM_GETSUBITEMRECT, properties.parentItemIndex.iItem, reinterpret_cast<LPARAM>(&rc))) {
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