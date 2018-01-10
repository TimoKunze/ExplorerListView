// ListViewFooterItems.cpp: Manages a collection of ListViewFooterItem objects

#include "stdafx.h"
#include "ListViewFooterItems.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewFooterItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewFooterItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListViewFooterItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListViewFooterItems>* pLvwFooterItemsObj = NULL;
	CComObject<ListViewFooterItems>::CreateInstance(&pLvwFooterItemsObj);
	pLvwFooterItemsObj->AddRef();

	// clone all settings
	pLvwFooterItemsObj->properties = properties;

	pLvwFooterItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLvwFooterItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ListViewFooterItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		LVFOOTERINFO footerInfo = {0};
		footerInfo.mask = LVFF_ITEMCOUNT;
		SendMessage(hWndLvw, LVM_GETFOOTERINFO, 0, reinterpret_cast<LPARAM>(&footerInfo));
		if(static_cast<UINT>(properties.nextEnumeratedFooterItem) >= footerInfo.cItems) {
			properties.nextEnumeratedFooterItem = 0;
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitFooterItem(properties.nextEnumeratedFooterItem, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal), FALSE);
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedFooterItem;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListViewFooterItems::Reset(void)
{
	properties.nextEnumeratedFooterItem = 0;
	return S_OK;
}

STDMETHODIMP ListViewFooterItems::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedFooterItem += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


ListViewFooterItems::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewFooterItems::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HRESULT ListViewFooterItems::Properties::WindowQueryInterface(REFIID requiredInterface, LPUNKNOWN* ppObject)
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


void ListViewFooterItems::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewFooterItems::get_Item(LONG footerItemIdentifier, FooterItemIdentifierTypeConstants footerItemIdentifierType/* = fiitIndex*/, IListViewFooterItem** ppFooterItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppFooterItem, IListViewFooterItem*);
	if(!ppFooterItem) {
		return E_POINTER;
	}

	// retrieve the footer item's zero-based index
	int itemIndex = -1;
	switch(footerItemIdentifierType) {
		case fiitIndex:
			itemIndex = footerItemIdentifier;
			break;
	}

	if(IsValidFooterItemIndex(itemIndex, properties.pOwnerExLvw)) {
		ClassFactory::InitFooterItem(itemIndex, properties.pOwnerExLvw, IID_IListViewFooterItem, reinterpret_cast<LPUNKNOWN*>(ppFooterItem), FALSE);
		return S_OK;
	}

	if(footerItemIdentifierType == fiitIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP ListViewFooterItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ListViewFooterItems::Add(BSTR footerItemText, LONG insertAt/* = -1*/, LONG iconIndex/* = 0*/, LONG itemData/* = 1*/, IListViewFooterItem** ppAddedFooterItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedFooterItem, IListViewFooterItem*);
	if(!ppAddedFooterItem) {
		return E_POINTER;
	}
	LONG footerItems = 0;
	Count(&footerItems);
	if(insertAt == -1) {
		insertAt = footerItems;
	}
	ATLASSERT(insertAt >= 0 && insertAt <= footerItems);
	if(insertAt < 0 || insertAt > footerItems) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	*ppAddedFooterItem = NULL;

	CComPtr<IListViewFooter> pListViewFooter = NULL;
	HRESULT hr = properties.WindowQueryInterface(IID_IListViewFooter, reinterpret_cast<LPUNKNOWN*>(&pListViewFooter));
	if(SUCCEEDED(hr)) {
		CW2W converter(footerItemText);
		hr = pListViewFooter->InsertButton(insertAt, converter, NULL, iconIndex, itemData);
		ClassFactory::InitFooterItem(insertAt, properties.pOwnerExLvw, IID_IListViewFooterItem, reinterpret_cast<LPUNKNOWN*>(ppAddedFooterItem));
	}
	return hr;
}

STDMETHODIMP ListViewFooterItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	LVFOOTERINFO footerInfo = {0};
	footerInfo.mask = LVFF_ITEMCOUNT;
	if(SendMessage(hWndLvw, LVM_GETFOOTERINFO, 0, reinterpret_cast<LPARAM>(&footerInfo))) {
		*pValue = footerInfo.cItems;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewFooterItems::RemoveAll(void)
{
	CComPtr<IListViewFooter> pListViewFooter = NULL;
	HRESULT hr = properties.WindowQueryInterface(IID_IListViewFooter, reinterpret_cast<LPUNKNOWN*>(&pListViewFooter));
	if(SUCCEEDED(hr)) {
		hr = pListViewFooter->RemoveAllButtons();
	}
	return hr;
}