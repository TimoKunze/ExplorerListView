// ListViewFooterItem.cpp: A wrapper for existing listview footer items.

#include "stdafx.h"
#include "ListViewFooterItem.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewFooterItem::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewFooterItem, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListViewFooterItem::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewFooterItem::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HRESULT ListViewFooterItem::Properties::WindowQueryInterface(REFIID requiredInterface, LPUNKNOWN* ppObject)
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


void ListViewFooterItem::Attach(int itemIndex)
{
	properties.itemIndex = itemIndex;
}

void ListViewFooterItem::Detach(void)
{
	properties.itemIndex = -1;
}

void ListViewFooterItem::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewFooterItem::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVFOOTERITEM footerItem = {0};
	footerItem.mask = LVFIF_STATE;
	footerItem.stateMask = LVFIS_FOCUSED;
	if(SendMessage(hWndLvw, LVM_GETFOOTERITEM, properties.itemIndex, reinterpret_cast<LPARAM>(&footerItem))) {
		*pValue = BOOL2VARIANTBOOL(footerItem.state & LVFIS_FOCUSED);;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewFooterItem::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemIndex;
	return S_OK;
}

STDMETHODIMP ListViewFooterItem::get_ItemData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	CComPtr<IListViewFooter> pListViewFooter = NULL;
	HRESULT hr = properties.WindowQueryInterface(IID_IListViewFooter, reinterpret_cast<LPUNKNOWN*>(&pListViewFooter));
	if(SUCCEEDED(hr)) {
		hr = pListViewFooter->GetButtonLParam(properties.itemIndex, pValue);
	}
	return hr;
}

STDMETHODIMP ListViewFooterItem::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVFOOTERITEM footerItem = {0};
	footerItem.mask = LVFIF_TEXT;
	footerItem.cchTextMax = MAX_FOOTERITEMTEXTLENGTH;
	footerItem.pszText = reinterpret_cast<LPWSTR>(malloc((footerItem.cchTextMax + 1) * sizeof(WCHAR)));
	if(SendMessage(hWndLvw, LVM_GETFOOTERITEM, properties.itemIndex, reinterpret_cast<LPARAM>(&footerItem))) {
		*pValue = _bstr_t(footerItem.pszText).Detach();
		SECUREFREE(footerItem.pszText);
		return S_OK;
	}
	SECUREFREE(footerItem.pszText);
	return E_FAIL;
}


STDMETHODIMP ListViewFooterItem::GetRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	RECT itemRectangle = {0};
	if(SendMessage(hWndLvw, LVM_GETFOOTERITEMRECT, properties.itemIndex, reinterpret_cast<LPARAM>(&itemRectangle))) {
		if(pXLeft) {
			*pXLeft = itemRectangle.left;
		}
		if(pYTop) {
			*pYTop = itemRectangle.top;
		}
		if(pXRight) {
			*pXRight = itemRectangle.right;
		}
		if(pYBottom) {
			*pYBottom = itemRectangle.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}