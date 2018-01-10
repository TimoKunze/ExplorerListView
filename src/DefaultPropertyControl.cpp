// DefaultPropertyControl.cpp: Default implementation of the IPropertyControl interface

#include "stdafx.h"
#include "DefaultPropertyControl.h"


#ifdef INCLUDESUBITEMCALLBACKCODE
	DefaultPropertyControl::Properties::~Properties()
	{
		if(pOwnerExLvw) {
			pOwnerExLvw->Release();
		}
	}


	//////////////////////////////////////////////////////////////////////
	// implementation of IPropertyControlBase
	STDMETHODIMP DefaultPropertyControl::Initialize(LPUNKNOWN, PROPDESC_CONTROL_TYPE)
	{
		ATLASSERT(FALSE && "Initialize");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::GetSize(PROPCTL_RECT_TYPE /*unknown1*/, HDC /*hDC*/, SIZE const* /*pUnknown2*/, LPSIZE /*pUnknown3*/)
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::SetWindowTheme(LPCWSTR, LPCWSTR)
	{
		ATLASSERT(FALSE && "SetWindowTheme");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::SetFont(HFONT hFont)
	{
		if(properties.editControl.IsWindow()) {
			properties.editControl.SetFont(hFont);
		}
		return S_OK;
	}

	STDMETHODIMP DefaultPropertyControl::SetTextColor(COLORREF /*color*/)
	{
		// must be implemented
		return S_OK;
	}

	STDMETHODIMP DefaultPropertyControl::GetFlags(PINT)
	{
		ATLASSERT(FALSE && "GetFlags");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::SetFlags(int /*flags*/, int /*mask*/)
	{
		// must be implemented
		return S_OK;
	}

	STDMETHODIMP DefaultPropertyControl::AdjustWindowRectPCB(HWND /*hWndListView*/, LPRECT /*pItemRectangle*/, RECT const* /*pUnknown1*/, int /*unknown2*/)
	{
		// must be implemented
		return S_OK;
	}

	STDMETHODIMP DefaultPropertyControl::SetValue(LPUNKNOWN)
	{
		ATLASSERT(FALSE && "SetValue");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::InvokeDefaultAction(void)
	{
		ATLASSERT(FALSE && "InvokeDefaultAction");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::Destroy(void)
	{
		if(properties.editControl.IsWindow()) {
			properties.editControl.DestroyWindow();
		}
		return S_OK;
	}

	STDMETHODIMP DefaultPropertyControl::SetFormatFlags(int)
	{
		ATLASSERT(FALSE && "SetFormatFlags");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::GetFormatFlags(PINT)
	{
		ATLASSERT(FALSE && "GetFormatFlags");
		return E_NOTIMPL;
	}
	// implementation of IPropertyControlBase
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	// implementation of IPropertyControl
	STDMETHODIMP DefaultPropertyControl::GetValue(REFIID /*requiredInterface*/, LPVOID* /*ppObject*/)
	{
		ATLASSERT(FALSE && "GetValue");
		return E_NOINTERFACE;
	}

	STDMETHODIMP DefaultPropertyControl::Create(HWND hWndParent, RECT const* pItemRectangle, RECT const* /*pUnknown1*/, int unknown2)
	{
		ATLASSERT(unknown2 == 0x02 || unknown2 == 0x0C);

		if(IsWindow(hWndParent)) {
			properties.editControl.Create(hWndParent, *const_cast<LPRECT>(pItemRectangle), NULL, WS_BORDER | WS_CHILDWINDOW | WS_VISIBLE | ES_AUTOHSCROLL | ES_WANTRETURN);
			properties.editControl.SetFont(CWindow(hWndParent).GetFont());
			properties.editControl.SetFocus();
			return S_OK;
		}
		return E_INVALIDARG;
	}

	STDMETHODIMP DefaultPropertyControl::SetPosition(RECT const*, RECT const*)
	{
		ATLASSERT(FALSE && "SetPosition");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::IsModified(LPBOOL /*pModified*/)
	{
		ATLASSERT(FALSE && "IsModified");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::SetModified(BOOL /*modified*/)
	{
		ATLASSERT(FALSE && "SetModified");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::ValidationFailed(LPCWSTR)
	{
		ATLASSERT(FALSE && "ValidationFailed");
		return E_NOTIMPL;
	}

	STDMETHODIMP DefaultPropertyControl::GetState(PINT /*pState*/)
	{
		ATLASSERT(FALSE && "GetState");
		return E_NOTIMPL;
	}
	// implementation of IPropertyControl
	//////////////////////////////////////////////////////////////////////


	void DefaultPropertyControl::SetOwner(__in_opt ExplorerListView* pOwner)
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