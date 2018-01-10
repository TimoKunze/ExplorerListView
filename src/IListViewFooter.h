//////////////////////////////////////////////////////////////////////
/// \class IListViewFooter
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for working with listview footers</em>
///
/// This interface is implemented by listview controls (comctl32.dll version 6.10 or higher). It is used to
/// configure footers.\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView, LVM_QUERYINTERFACE,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/comctl32/controls/listview/interfaces/ilistviewfooter.htm">IListViewFooter</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IListViewFooterCallback.h"


// the interface's GUID
const IID IID_IListViewFooter = {0xF0034DA8, 0x8A22, 0x4151, {0x8F, 0x16, 0x2E, 0xBA, 0x76, 0x56, 0x5B, 0xCC}};


class IListViewFooter :
    public IUnknown
{
public:
	/// \brief <em>Retrieves whether the footer area is currently displayed</em>
	///
	/// Retrieves whether the listview control's footer area is currently displayed.
	///
	/// \param[out] pVisible \c TRUE if the footer area is visible; otherwise \c FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Show, ExplorerListView::IsFooterVisible
	virtual HRESULT STDMETHODCALLTYPE IsVisible(PINT pVisible) = 0;
	/// \brief <em>Retrieves the caret footer item</em>
	///
	/// Retrieves the listview control's focused footer item.
	///
	/// \param[out] pItemIndex Receives the zero-based index of the footer item that has the keyboard focus.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetFooterFocus, ExplorerListView::get_CaretFooterItem
	virtual HRESULT STDMETHODCALLTYPE GetFooterFocus(PINT pItemIndex) = 0;
	/// \brief <em>Sets the caret footer item</em>
	///
	/// Sets the listview control's focused footer item.
	///
	/// \param[in] itemIndex The zero-based index of the footer item to which to set the keyboard focus.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetFooterFocus, ExplorerListView::putref_CaretFooterItem
	virtual HRESULT STDMETHODCALLTYPE SetFooterFocus(int itemIndex) = 0;
	/// \brief <em>Sets the footer area's caption</em>
	///
	/// Sets the title text of the listview control's footer area.
	///
	/// \param[in] pText The text to display in the footer area's title.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa InsertButton, ExplorerListView::put_FooterIntroText
	virtual HRESULT STDMETHODCALLTYPE SetIntroText(LPCWSTR pText) = 0;
	/// \brief <em>Makes the footer area visible</em>
	///
	/// Makes the listview control's footer area visible and registers the callback object that is notified
	/// about item clicks and item deletions.
	///
	/// \param[in] pCallbackObject The \c IListViewFooterCallback implementation of the callback object to
	///            register.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IListViewFooterCallback, IsVisible, ExplorerListView::ShowFooter
	virtual HRESULT STDMETHODCALLTYPE Show(IListViewFooterCallback* pCallbackObject) = 0;
	/// \brief <em>Removes all footer items</em>
	///
	/// Removes all footer items from the listview control's footer area.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa InsertButton, ListViewFooterItems::RemoveAll
	virtual HRESULT STDMETHODCALLTYPE RemoveAllButtons(void) = 0;
	/// \brief <em>Inserts a footer item</em>
	///
	/// Inserts a new footer item with the specified properties at the specified position into the listview
	/// control.
	///
	/// \param[in] insertAt The zero-based index at which to insert the new footer item.
	/// \param[in] pText The new footer item's text.
	/// \param[in] pUnknown TODO
	/// \param[in] iconIndex The zero-based index of the new footer item's icon.
	/// \param[in] lParam The integer data that will be associated with the new footer item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa RemoveAllButtons, ListViewFooterItems::Add
	virtual HRESULT STDMETHODCALLTYPE InsertButton(int insertAt, LPCWSTR pText, LPCWSTR pUnknown, UINT iconIndex, LONG lParam) = 0;
	/// \brief <em>Retrieves a footer item's associated data</em>
	///
	/// Retrieves the integer data associated with the specified footer item.
	///
	/// \param[in] itemIndex The zero-based index of the footer for which to retrieve the associated data.
	/// \param[out] pLParam Receives the associated data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa InsertButton, ListViewFooterItem::get_ItemData
	virtual HRESULT STDMETHODCALLTYPE GetButtonLParam(int itemIndex, LONG* pLParam) = 0;
};