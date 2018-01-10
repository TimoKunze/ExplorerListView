// ListViewWorkAreas.cpp: Manages a collection of ListViewWorkArea objects

#include "stdafx.h"
#include "ListViewWorkAreas.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewWorkAreas::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewWorkAreas, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListViewWorkAreas::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListViewWorkAreas>* pLvwWorkAreasObj = NULL;
	CComObject<ListViewWorkAreas>::CreateInstance(&pLvwWorkAreasObj);
	pLvwWorkAreasObj->AddRef();

	// clone all settings
	pLvwWorkAreasObj->properties = properties;

	pLvwWorkAreasObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLvwWorkAreasObj->Release();
	return S_OK;
}

STDMETHODIMP ListViewWorkAreas::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(pItems == NULL) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	int workAreas = 0;
	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(&workAreas));

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		if(properties.nextEnumeratedWorkArea >= workAreas) {
			properties.nextEnumeratedWorkArea = 0;
			// there's nothing more to iterate
			break;
		}

		ClassFactory::InitWorkArea(properties.nextEnumeratedWorkArea, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal), FALSE);
		pItems[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedWorkArea;
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListViewWorkAreas::Reset(void)
{
	properties.nextEnumeratedWorkArea = 0;
	return S_OK;
}

STDMETHODIMP ListViewWorkAreas::Skip(ULONG numberOfItemsToSkip)
{
	properties.nextEnumeratedWorkArea += numberOfItemsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


ListViewWorkAreas::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewWorkAreas::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListViewWorkAreas::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewWorkAreas::get_Item(LONG workAreaIndex, IListViewWorkArea** ppWorkArea)
{
	ATLASSERT_POINTER(ppWorkArea, IListViewWorkArea*);
	if(!ppWorkArea) {
		return E_POINTER;
	}

	if(IsValidWorkAreaIndex(workAreaIndex, properties.pOwnerExLvw)) {
		ClassFactory::InitWorkArea(workAreaIndex, properties.pOwnerExLvw, IID_IListViewWorkArea, reinterpret_cast<LPUNKNOWN*>(ppWorkArea), FALSE);
		return S_OK;
	}
	return DISP_E_BADINDEX;
}

STDMETHODIMP ListViewWorkAreas::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ListViewWorkAreas::Add(OLE_XPOS_PIXELS xLeft, OLE_YPOS_PIXELS yTop, OLE_XPOS_PIXELS xRight, OLE_YPOS_PIXELS yBottom, LONG insertAt/* = -1*/, IListViewWorkArea** ppAddedWorkArea/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedWorkArea, IListViewWorkArea*);
	if(!ppAddedWorkArea) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	int workAreas = 0;
	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(&workAreas));
	if(workAreas == LV_MAX_WORKAREAS) {
		return E_FAIL;
	}
	if(insertAt == -1) {
		insertAt = workAreas;
	}
	ATLASSERT((insertAt >= 0) && (insertAt <= workAreas));
	if((insertAt < 0) || (insertAt > workAreas)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	LPRECT pWorkAreaRectangles = reinterpret_cast<LPRECT>(HeapAlloc(GetProcessHeap(), 0, (workAreas + 1) * sizeof(RECT)));
	if(!pWorkAreaRectangles) {
		return E_OUTOFMEMORY;
	}
	if(workAreas > 0) {
		LPRECT pCurrentWorkAreaRectangles = reinterpret_cast<LPRECT>(HeapAlloc(GetProcessHeap(), 0, workAreas * sizeof(RECT)));
		if(!pCurrentWorkAreaRectangles) {
			HeapFree(GetProcessHeap(), 0, pWorkAreaRectangles);
			return E_OUTOFMEMORY;
		}
		SendMessage(hWndLvw, LVM_GETWORKAREAS, workAreas, reinterpret_cast<LPARAM>(pCurrentWorkAreaRectangles));
		if(insertAt > 0) {
			CopyMemory(pWorkAreaRectangles, pCurrentWorkAreaRectangles, insertAt * sizeof(RECT));
		}
		if(insertAt < workAreas) {
			CopyMemory(pWorkAreaRectangles + insertAt + 1, pCurrentWorkAreaRectangles + insertAt, (workAreas - insertAt) * sizeof(RECT));
		}
		HeapFree(GetProcessHeap(), 0, pCurrentWorkAreaRectangles);
	}
	pWorkAreaRectangles[insertAt].left = xLeft;
	pWorkAreaRectangles[insertAt].top = yTop;
	pWorkAreaRectangles[insertAt].right = xRight;
	pWorkAreaRectangles[insertAt].bottom = yBottom;

	// finally insert the working area
	SendMessage(hWndLvw, LVM_SETWORKAREAS, workAreas + 1, reinterpret_cast<LPARAM>(pWorkAreaRectangles));
	ClassFactory::InitWorkArea(insertAt, properties.pOwnerExLvw, IID_IListViewWorkArea, reinterpret_cast<LPUNKNOWN*>(ppAddedWorkArea));

	HeapFree(GetProcessHeap(), 0, pWorkAreaRectangles);
	return S_OK;
}

STDMETHODIMP ListViewWorkAreas::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(pValue));
	return S_OK;
}

STDMETHODIMP ListViewWorkAreas::Remove(LONG workAreaIndex)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	int workAreas = 0;
	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(&workAreas));
	if((workAreaIndex >= 0) && (workAreaIndex < workAreas)) {
		LPRECT pWorkAreaRectangles = new RECT[workAreas];
		SendMessage(hWndLvw, LVM_GETWORKAREAS, workAreas, reinterpret_cast<LPARAM>(pWorkAreaRectangles));

		if(workAreaIndex < workAreas - 1) {
			CopyMemory(pWorkAreaRectangles + workAreaIndex, pWorkAreaRectangles + workAreaIndex + 1, (workAreas - workAreaIndex - 1) * sizeof(RECT));
		}

		SendMessage(hWndLvw, LVM_SETWORKAREAS, workAreas - 1, reinterpret_cast<LPARAM>(pWorkAreaRectangles));

		delete[] pWorkAreaRectangles;
		return S_OK;
	}

	return DISP_E_BADINDEX;
}

STDMETHODIMP ListViewWorkAreas::RemoveAll(void)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	SendMessage(hWndLvw, LVM_SETWORKAREAS, 0, 0);
	return S_OK;
}