// VirtualListViewWorkArea.cpp: A wrapper for non-existing listview working areas.

#include "stdafx.h"
#include "VirtualListViewWorkArea.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualListViewWorkArea::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualListViewWorkArea, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


void VirtualListViewWorkArea::Attach(LPRECT pWorkAreas, int workAreaIndex)
{
	properties.pWorkAreas = pWorkAreas;
	properties.workAreaIndex = workAreaIndex;
}

void VirtualListViewWorkArea::Detach(void)
{
	properties.pWorkAreas = NULL;
	properties.workAreaIndex = -1;
}


STDMETHODIMP VirtualListViewWorkArea::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.workAreaIndex;
	return S_OK;
}


STDMETHODIMP VirtualListViewWorkArea::GetRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(pXLeft) {
		*pXLeft = properties.pWorkAreas[properties.workAreaIndex].left;
	}
	if(pYTop) {
		*pYTop = properties.pWorkAreas[properties.workAreaIndex].top;
	}
	if(pXRight) {
		*pXRight = properties.pWorkAreas[properties.workAreaIndex].right;
	}
	if(pYBottom) {
		*pYBottom = properties.pWorkAreas[properties.workAreaIndex].bottom;
	}
	return S_OK;
}