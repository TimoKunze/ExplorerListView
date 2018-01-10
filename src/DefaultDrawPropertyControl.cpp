// DefaultDrawPropertyControl.cpp: Default implementation of the IDrawPropertyControl interface

#include "stdafx.h"
#include "DefaultDrawPropertyControl.h"


#ifdef INCLUDESUBITEMCALLBACKCODE
	DefaultDrawPropertyControl::Properties::~Properties()
	{
		if(pOwnerExLvw) {
			pOwnerExLvw->Release();
		}
	}


	//////////////////////////////////////////////////////////////////////
	// implementation of IPropertyControlBase
	STDMETHODIMP DefaultDrawPropertyControl::Initialize(LPUNKNOWN, PROPDESC_CONTROL_TYPE)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::GetSize(PROPCTL_RECT_TYPE /*unknown1*/, HDC /*hDC*/, SIZE const* /*pUnknown2*/, LPSIZE /*pUnknown3*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::SetWindowTheme(LPCWSTR, LPCWSTR)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::SetFont(HFONT /*hFont*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::SetTextColor(COLORREF /*color*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::GetFlags(PINT)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::SetFlags(int /*flags*/, int /*mask*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::AdjustWindowRectPCB(HWND /*hWndListView*/, LPRECT /*pItemRectangle*/, RECT const* /*pUnknown1*/, int /*unknown2*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::SetValue(LPUNKNOWN)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::InvokeDefaultAction(void)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::Destroy(void)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::SetFormatFlags(int)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::GetFormatFlags(PINT)
	{
		return E_NOTIMPL;
	}
	// implementation of IPropertyControlBase
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	// implementation of IDrawPropertyControl
	STDMETHODIMP DefaultDrawPropertyControl::GetDrawFlags(PINT)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::WindowlessDraw(HDC /*hDC*/, RECT const* /*pDrawingRectangle*/, int /*unknown*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::HasVisibleContent(void)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::GetDisplayText(LPWSTR*)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultDrawPropertyControl::GetTooltipInfo(HDC, SIZE const*, PINT)
	{
		return E_NOTIMPL;
	}
	// implementation of IDrawPropertyControl
	//////////////////////////////////////////////////////////////////////


	void DefaultDrawPropertyControl::SetOwner(__in_opt ExplorerListView* pOwner)
	{
		if(properties.pOwnerExLvw) {
			properties.pOwnerExLvw->Release();
		}
		properties.pOwnerExLvw = pOwner;
		if(properties.pOwnerExLvw) {
			properties.pOwnerExLvw->AddRef();
		}
	}
#endif