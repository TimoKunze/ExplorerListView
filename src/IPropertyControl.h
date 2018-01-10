//////////////////////////////////////////////////////////////////////
/// \class IPropertyControl
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for supporting enhanced label-editing</em>
///
/// This interface is implemented by client applications and used by listview controls (comctl32.dll
/// version 6.10 or higher) to provide enhanced label-editing features. By implementing this interface,
/// the following features become possible:
/// - edit listview group header labels
/// - edit sub-item labels
/// - provide other control's than a text box for label-editing (it should also be possible to edit labels
///   without any control)
/// - start label-editing mode on mouse hover\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView, IPropertyControlBase, DefaultPropertyControl,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/shell32/interfaces/ipropertycontrol.htm">IPropertyControl</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IPropertyControlBase.h"


// the interface's GUID
const IID IID_IPropertyControl = {0x5E82A4DD, 0x9561, 0x476A, {0x86, 0x34, 0x1B, 0xEB, 0xAC, 0xBA, 0x4A, 0x38}};


class IPropertyControl :
    public IPropertyControlBase
{
public:
	virtual HRESULT STDMETHODCALLTYPE GetValue(REFIID requiredInterface, LPVOID* ppObject) = 0;
	// possible values for unknown2 (list may be incomplete):
	//   0x02 - mouse has been moved over the sub-item
	//   0x0C - ISubItemCallBack::EditSubItem has been called or a slow double click has been made
	virtual HRESULT STDMETHODCALLTYPE Create(HWND hWndParent, RECT const* pItemRectangle, RECT const* pUnknown1, int unknown2) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetPosition(RECT const*, RECT const*) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsModified(LPBOOL pModified) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetModified(BOOL modified) = 0;
	virtual HRESULT STDMETHODCALLTYPE ValidationFailed(LPCWSTR) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetState(PINT pState) = 0;
};