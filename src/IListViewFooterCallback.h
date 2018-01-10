//////////////////////////////////////////////////////////////////////
/// \class IListViewFooterCallback
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for working with listview footers</em>
///
/// This interface is implemented by client applications and used by listview controls (comctl32.dll
/// version 6.10 or higher) to notify the client application about footer button clicks.\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation. It's not fully clear what the last parameter of the \c OnButtonClicked
///       method is used for.
///
/// \sa ExplorerListView, IListViewFooter,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/comctl32/controls/listview/interfaces/ilistviewfootercallback.htm">IListViewFooterCallback</a>
//////////////////////////////////////////////////////////////////////


#pragma once


const IID IID_IListViewFooterCallback = {0x88EB9442, 0x913B, 0x4AB4, {0xA7, 0x41, 0xDD, 0x99, 0xDC, 0xB7, 0x55, 0x8B}};


class IListViewFooterCallback :
    public IUnknown
{
public:
	/// \brief <em>Notifies the client that a footer item has been clicked</em>
	///
	/// This method is called by the listview control to notify the client application that the user has
	/// clicked a footer item.
	///
	/// \param[in] itemIndex The zero-based index of the footer item that has been clicked.
	/// \param[in] lParam The application-defined integer value that is associated with the clicked button.
	/// \param[out] pRemoveFooter If set to \c TRUE, the listview control will remove the footer area.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::Raise_FooterItemClick
	virtual HRESULT STDMETHODCALLTYPE OnButtonClicked(int itemIndex, LPARAM lParam, PINT pRemoveFooter) = 0;
	/// \brief <em>Notifies the client that a footer item has been removed</em>
	///
	/// This method is called by the listview control to notify the client application that it has removed a
	/// footer item.
	///
	/// \param[in] itemIndex The zero-based index of the footer item that has been removed.
	/// \param[in] lParam The application-defined integer value that is associated with the removed button.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::Raise_FreeFooterItemData
	virtual HRESULT STDMETHODCALLTYPE OnDestroyButton(int itemIndex, LPARAM lParam) = 0;
};