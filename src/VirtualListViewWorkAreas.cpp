// VirtualListViewWorkAreas.cpp: Manages a collection of VirtualListViewWorkArea objects

#include "stdafx.h"
#include "VirtualListViewWorkAreas.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualListViewWorkAreas::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualListViewWorkAreas, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP VirtualListViewWorkAreas::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<VirtualListViewWorkAreas>* pVLvwWorkAreasObj = NULL;
	CComObject<VirtualListViewWorkAreas>::CreateInstance(&pVLvwWorkAreasObj);
	pVLvwWorkAreasObj->AddRef();

	// clone all settings
	pVLvwWorkAreasObj->properties = properties;

	pVLvwWorkAreasObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pVLvwWorkAreasObj->Release();
	return S_OK;
}

STDMETHODIMP VirtualListViewWorkAreas::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		if(properties.nextEnumeratedWorkArea >= properties.numberOfWorkAreas) {
			properties.nextEnumeratedWorkArea = 0;
			// there's nothing more to iterate
			break;
		}

		pItems[i].pdispVal = NULL;

		// create and initialize a VirtualListViewWorkArea object
		CComObject<VirtualListViewWorkArea>* pVLvwWorkAreaObj = NULL;
		CComObject<VirtualListViewWorkArea>::CreateInstance(&pVLvwWorkAreaObj);
		pVLvwWorkAreaObj->AddRef();
		pVLvwWorkAreaObj->Attach(properties.pWorkAreas, properties.nextEnumeratedWorkArea);
		pVLvwWorkAreaObj->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
		pVLvwWorkAreaObj->Release();

		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedWorkArea;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP VirtualListViewWorkAreas::Reset(void)
{
	properties.nextEnumeratedWorkArea = 0;
	return S_OK;
}

STDMETHODIMP VirtualListViewWorkAreas::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedWorkArea += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


void VirtualListViewWorkAreas::Attach(int numberOfWorkAreas, LPRECT pWorkAreas)
{
	properties.numberOfWorkAreas = numberOfWorkAreas;
	properties.pWorkAreas = pWorkAreas;
}

void VirtualListViewWorkAreas::Detach(void)
{
	properties.numberOfWorkAreas = 0;
	if(properties.pWorkAreas) {
		delete[] properties.pWorkAreas;
		properties.pWorkAreas = NULL;
	}
}


STDMETHODIMP VirtualListViewWorkAreas::get_Item(LONG workAreaIndex, IVirtualListViewWorkArea** ppWorkArea)
{
	ATLASSERT_POINTER(ppWorkArea, IVirtualListViewWorkArea*);
	if(!ppWorkArea) {
		return E_POINTER;
	}

	if((workAreaIndex >= 0) && (workAreaIndex < properties.numberOfWorkAreas)) {
		*ppWorkArea = NULL;

		// create and initialize a VirtualListViewWorkArea object
		CComObject<VirtualListViewWorkArea>* pVLvwWorkAreaObj = NULL;
		CComObject<VirtualListViewWorkArea>::CreateInstance(&pVLvwWorkAreaObj);
		pVLvwWorkAreaObj->AddRef();
		pVLvwWorkAreaObj->Attach(properties.pWorkAreas, workAreaIndex);
		pVLvwWorkAreaObj->QueryInterface(IID_IVirtualListViewWorkArea, reinterpret_cast<LPVOID*>(ppWorkArea));
		pVLvwWorkAreaObj->Release();

		return S_OK;
	}
	return DISP_E_BADINDEX;
}

STDMETHODIMP VirtualListViewWorkAreas::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP VirtualListViewWorkAreas::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.numberOfWorkAreas;
	return S_OK;
}