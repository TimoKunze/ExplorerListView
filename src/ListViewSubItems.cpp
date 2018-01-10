// ListViewSubItems.cpp: Manages a collection of ListViewSubItem objects

#include "stdafx.h"
#include "ListViewSubItems.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewSubItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewSubItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListViewSubItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListViewSubItems>* pLvwSubItemsObj = NULL;
	CComObject<ListViewSubItems>::CreateInstance(&pLvwSubItemsObj);
	pLvwSubItemsObj->AddRef();

	// clone all settings
	pLvwSubItemsObj->properties = properties;

	pLvwSubItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLvwSubItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ListViewSubItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndHeader = properties.GetHeaderHWnd();
	ATLASSERT(IsWindow(hWndHeader));

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		if(properties.nextEnumeratedSubItem >= static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0))) {
			properties.nextEnumeratedSubItem = 0;
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitListSubItem(properties.parentItemIndex, properties.nextEnumeratedSubItem, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal), FALSE, FALSE);
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedSubItem;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListViewSubItems::Reset(void)
{
	properties.nextEnumeratedSubItem = 1;
	return S_OK;
}

STDMETHODIMP ListViewSubItems::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedSubItem += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


ListViewSubItems::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewSubItems::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}

HWND ListViewSubItems::Properties::GetHeaderHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWndHeader(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListViewSubItems::Attach(LVITEMINDEX& parentItemIndex)
{
	properties.parentItemIndex = parentItemIndex;
}

void ListViewSubItems::Detach(void)
{
	properties.parentItemIndex.iItem = -1;
	properties.parentItemIndex.iGroup = 0;
}

void ListViewSubItems::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewSubItems::get_Item(LONG subItemIdentifier, ColumnIdentifierTypeConstants subItemIdentifierType/* = citIndex*/, IListViewSubItem** ppSubItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppSubItem, IListViewSubItem*);
	if(!ppSubItem) {
		return E_POINTER;
	}

	// get the sub-item's index
	int subItemIndex = -1;
	switch(subItemIdentifierType) {
		case citIndex:
			if(subItemIdentifier >= 0) {
				HWND hWndHeader = properties.GetHeaderHWnd();
				ATLASSERT(IsWindow(hWndHeader));

				int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
				if(columnCount == -1) {
					return E_FAIL;
				}
				if(subItemIdentifier < columnCount) {
					subItemIndex = subItemIdentifier;
				}
			}
			break;
		case citPosition:
			if(subItemIdentifier >= 0) {
				HWND hWndHeader = properties.GetHeaderHWnd();
				ATLASSERT(IsWindow(hWndHeader));

				int columnCount = static_cast<int>(SendMessage(hWndHeader, HDM_GETITEMCOUNT, 0, 0));
				if(columnCount == -1) {
					return E_FAIL;
				}
				if(subItemIdentifier < columnCount) {
					subItemIndex = static_cast<int>(SendMessage(hWndHeader, HDM_ORDERTOINDEX, subItemIdentifier, 0));
				}
			}
			break;
		case citID:
			if(properties.pOwnerExLvw) {
				subItemIndex = properties.pOwnerExLvw->IDToColumnIndex(subItemIdentifier);
			}
			break;
	}

	ClassFactory::InitListSubItem(properties.parentItemIndex, subItemIndex, properties.pOwnerExLvw, IID_IListViewSubItem, reinterpret_cast<LPUNKNOWN*>(ppSubItem), FALSE);
	if(!*ppSubItem) {
		if(subItemIdentifierType == citIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}

	return S_OK;
}

STDMETHODIMP ListViewSubItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}

STDMETHODIMP ListViewSubItems::get_ParentItem(IListViewItem** ppParentItem)
{
	ATLASSERT_POINTER(ppParentItem, IListViewItem*);
	if(!ppParentItem) {
		return E_POINTER;
	}

	ClassFactory::InitListItem(properties.parentItemIndex, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppParentItem));
	return S_OK;
}


STDMETHODIMP ListViewSubItems::Count(LONG* pValue)
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
	} else {
		// we've 1 sub-item less than columns
		--(*pValue);
	}

	return S_OK;
}