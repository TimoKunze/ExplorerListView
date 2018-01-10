// ListViewWorkArea.cpp: A wrapper for existing listview working areas.

#include "stdafx.h"
#include "ListViewWorkArea.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewWorkArea::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewWorkArea, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ListViewWorkArea::Properties::~Properties()
{
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

HWND ListViewWorkArea::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListViewWorkArea::Attach(int workAreaIndex)
{
	properties.workAreaIndex = workAreaIndex;
}

void ListViewWorkArea::Detach(void)
{
	properties.workAreaIndex = -1;
}

void ListViewWorkArea::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewWorkArea::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.workAreaIndex;
	return S_OK;
}

STDMETHODIMP ListViewWorkArea::get_ListItems(IListViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IListViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitListItems(properties.pOwnerExLvw, IID_IListViewItems, reinterpret_cast<LPUNKNOWN*>(ppItems));

	if(*ppItems) {
		// install a fpWorkArea filter
		VARIANT filter;
		VariantInit(&filter);
		// create the array
		filter.vt = VT_ARRAY | VT_VARIANT;
		filter.parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		// store properties.workAreaIndex in the array (as object)
		VARIANT element;
		VariantInit(&element);
		element.vt = VT_DISPATCH;
		ClassFactory::InitWorkArea(properties.workAreaIndex, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&element.pdispVal));
		LONG index = 1;
		SafeArrayPutElement(filter.parray, &index, &element);
		VariantClear(&element);

		(*ppItems)->put_FilterType(fpWorkArea, ftIncluding);
		(*ppItems)->put_Filter(fpWorkArea, filter);
		VariantClear(&filter);
	}

	return S_OK;
}


STDMETHODIMP ListViewWorkArea::GetRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	int workAreas = 0;
	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(&workAreas));
	if(workAreas > 0) {
		LPRECT pWorkAreaRectangles = new RECT[workAreas];
		SendMessage(hWndLvw, LVM_GETWORKAREAS, workAreas, reinterpret_cast<LPARAM>(pWorkAreaRectangles));

		if(pXLeft) {
			*pXLeft = pWorkAreaRectangles[properties.workAreaIndex].left;
		}
		if(pYTop) {
			*pYTop = pWorkAreaRectangles[properties.workAreaIndex].top;
		}
		if(pXRight) {
			*pXRight = pWorkAreaRectangles[properties.workAreaIndex].right;
		}
		if(pYBottom) {
			*pYBottom = pWorkAreaRectangles[properties.workAreaIndex].bottom;
		}

		delete[] pWorkAreaRectangles;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ListViewWorkArea::SetRectangle(OLE_XPOS_PIXELS xLeft, OLE_YPOS_PIXELS yTop, OLE_XPOS_PIXELS xRight, OLE_YPOS_PIXELS yBottom)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	int workAreas = 0;
	SendMessage(hWndLvw, LVM_GETNUMBEROFWORKAREAS, 0, reinterpret_cast<LPARAM>(&workAreas));
	if(workAreas > 0) {
		LPRECT pWorkAreaRectangles = new RECT[workAreas];
		SendMessage(hWndLvw, LVM_GETWORKAREAS, workAreas, reinterpret_cast<LPARAM>(pWorkAreaRectangles));

		pWorkAreaRectangles[properties.workAreaIndex].left = xLeft;
		pWorkAreaRectangles[properties.workAreaIndex].top = yTop;
		pWorkAreaRectangles[properties.workAreaIndex].right = xRight;
		pWorkAreaRectangles[properties.workAreaIndex].bottom = yBottom;

		SendMessage(hWndLvw, LVM_SETWORKAREAS, workAreas, reinterpret_cast<LPARAM>(pWorkAreaRectangles));

		delete[] pWorkAreaRectangles;
		return S_OK;
	}
	return E_FAIL;
}