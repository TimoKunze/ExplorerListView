//////////////////////////////////////////////////////////////////////
/// \class IDrawPropertyControl
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>TODO</em>
///
/// TODO\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView, ISubItemCallback::GetSubItemControl, DefaultDrawPropertyControl,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/shell32/interfaces/idrawpropertycontrol.htm">IDrawPropertyControl</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IPropertyControlBase.h"


// the interface's GUID
const IID IID_IDrawPropertyControl = {0xE6DFF6FD, 0xBCD5, 0x4162, {0x9C, 0x65, 0xA3, 0xB1, 0x8C, 0x61, 0x6F, 0xDB}};
//{1572DD51-443C-44B0-ACE4-38A005FC697E}
const IID IID_IWin10Unknown = {0x1572DD51, 0x443C, 0x44B0, {0xAC, 0xE4, 0x38, 0xA0, 0x05, 0xFC, 0x69, 0x7E}};


class IDrawPropertyControl :
    public IPropertyControlBase
{
public:
	// All parameter names have been guessed!
	virtual HRESULT STDMETHODCALLTYPE GetDrawFlags(PINT) = 0;
	virtual HRESULT STDMETHODCALLTYPE WindowlessDraw(HDC hDC, RECT const* pDrawingRectangle, int a) = 0;
	virtual HRESULT STDMETHODCALLTYPE HasVisibleContent(void) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetDisplayText(LPWSTR*) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetTooltipInfo(HDC, SIZE const*, PINT) = 0;
};