//////////////////////////////////////////////////////////////////////
/// \class IPropertyControlBase
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>TODO</em>
///
/// TODO\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView, IPropertyControl,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/shell32/interfaces/ipropertycontrolbase.htm">IPropertyControlBase</a>
//////////////////////////////////////////////////////////////////////


#pragma once


// the interface's GUID
const IID IID_IPropertyControlBase = {0x6E71A510, 0x732A, 0x4557, {0x95, 0x96, 0xA8, 0x27, 0xE3, 0x6D, 0xAF, 0x8F}};


class IPropertyControlBase :
    public IUnknown
{
public:
	// All parameter names have been guessed!
	// 2 -> Calendar control becomes a real calendar control instead of a date picker (but with a height of only 1 line)
	// 3 -> Calendar control becomes a simple text box
	typedef enum tagPROPDESC_CONTROL_TYPE PROPDESC_CONTROL_TYPE;
	typedef enum tagPROPCTL_RECT_TYPE PROPCTL_RECT_TYPE;
	virtual HRESULT STDMETHODCALLTYPE Initialize(LPUNKNOWN, PROPDESC_CONTROL_TYPE) = 0;
	// called when editing group labels
	// unknown1 might be a LVGGR_* value
	// hDC seems to be always NULL
	// pUnknown2 seems to be always NULL
	// pUnknown3 seems to receive the size set by the method
	virtual HRESULT STDMETHODCALLTYPE GetSize(PROPCTL_RECT_TYPE unknown1, HDC hDC, SIZE const* pUnknown2, LPSIZE pUnknown3) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetWindowTheme(LPCWSTR, LPCWSTR) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetFont(HFONT hFont) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetTextColor(COLORREF color) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFlags(PINT) = 0;
	// IPropertyControl:
	//   called before the edit control is created and before it is dismissed
	//   flags is 1 on first call and 0 on second call (mask is always 1)
	virtual HRESULT STDMETHODCALLTYPE SetFlags(int flags, int mask) = 0;
	// possible values for unknown2 (list may be incomplete):
	//   0x02 - mouse has been moved over the sub-item
	//   0x0C - ISubItemCallBack::EditSubItem has been called or a slow double click has been made
	virtual HRESULT STDMETHODCALLTYPE AdjustWindowRectPCB(HWND hWndListView, LPRECT pItemRectangle, RECT const* pUnknown1, int unknown2) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetValue(LPUNKNOWN) = 0;
	virtual HRESULT STDMETHODCALLTYPE InvokeDefaultAction(void) = 0;
	virtual HRESULT STDMETHODCALLTYPE Destroy(void) = 0;
	// 0x01 - property description will be displayed in front of the property value
	virtual HRESULT STDMETHODCALLTYPE SetFormatFlags(int) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFormatFlags(PINT) = 0;
};