//////////////////////////////////////////////////////////////////////
/// \class IOwnerDataCallback
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for supporting enhanced virtual mode</em>
///
/// This interface is implemented by client applications and used by listview controls (comctl32.dll
/// version 6.10 or higher) to provide enhanced virtual mode features. By implementing this interface,
/// the following virtual mode features become possible:
/// - listview groups
/// - multiple occurrences of the same item (each of them may be in different groups)
/// - free item positioning\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView, IListView_WINVISTA, IListView_WIN7,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/comctl32/controls/listview/interfaces/iownerdatacallback.htm">IOwnerDataCallback</a>
//////////////////////////////////////////////////////////////////////


#pragma once


// the interface's GUID
const IID IID_IOwnerDataCallback = {0x44C09D56, 0x8D3B, 0x419D, {0xA4, 0x62, 0x7B, 0x95, 0x6B, 0x10, 0x5B, 0x47}};


class IOwnerDataCallback :
    public IUnknown
{
public:
	// All parameter names have been guessed!
	/// \brief <em>TODO</em>
	///
	/// TODO
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetItemPosition
	virtual HRESULT STDMETHODCALLTYPE GetItemPosition(int itemIndex, LPPOINT pPosition) = 0;
	/// \brief <em>TODO</em>
	///
	/// TODO
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetItemPosition
	virtual HRESULT STDMETHODCALLTYPE SetItemPosition(int itemIndex, POINT position) = 0;
	/// \brief <em>Will be called to retrieve an item's zero-based control-wide index</em>
	///
	/// This method is called by the listview control to retrieve an item's zero-based control-wide index.
	/// The item is identified by a zero-based group index, which identifies the listview group in which
	/// the item is displayed, and a zero-based group-wide item index, which identifies the item within its
	/// group.
	///
	/// \param[in] groupIndex The zero-based index of the listview group containing the item.
	/// \param[in] groupWideItemIndex The item's zero-based group-wide index within the listview group
	///            specified by \c groupIndex.
	/// \param[out] pTotalItemIndex Receives the item's zero-based control-wide index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::Raise_MapGroupWideToTotalItemIndex
	virtual HRESULT STDMETHODCALLTYPE GetItemInGroup(int groupIndex, int groupWideItemIndex, PINT pTotalItemIndex) = 0;
	/// \brief <em>Will be called to retrieve the group containing a specific occurrence of an item</em>
	///
	/// This method is called by the listview control to retrieve the listview group in which the specified
	/// occurrence of the specified item is displayed.
	///
	/// \param[in] itemIndex The item's zero-based (control-wide) index.
	/// \param[in] occurrenceIndex The zero-based index of the item's copy for which the group membership is
	///            retrieved.
	/// \param[out] pGroupIndex Receives the zero-based index of the listview group that shall contain the
	///             specified copy of the specified item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::Raise_ItemGetGroup, GetItemGroupCount
	virtual HRESULT STDMETHODCALLTYPE GetItemGroup(int itemIndex, int occurenceIndex, PINT pGroupIndex) = 0;
	/// \brief <em>Will be called to determine how often an item occurs in the listview control</em>
	///
	/// This method is called by the listview control to determine how often the specified item occurs in the
	/// listview control.
	///
	/// \param[in] itemIndex The item's zero-based (control-wide) index.
	/// \param[out] pOccurrencesCount Receives the number of occurrences of the item in the listview control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::Raise_ItemGetOccurrencesCount, GetItemGroup
	virtual HRESULT STDMETHODCALLTYPE GetItemGroupCount(int itemIndex, PINT pOccurenceCount) = 0;
	/// \brief <em>Will be called to prepare the client app that the data for a certain range of items will be required very soon</em>
	///
	/// This method is similar to the \c LVN_ODCACHEHINT notification. It tells the client application that
	/// it should preload the details for a certain range of items because the listview control is about to
	/// request these details. The difference to \c LVN_ODCACHEHINT is that this method identifies the items
	/// by their zero-based group-wide index and the zero-based index of the listview group containing the
	/// item.
	///
	/// \param[in] firstItem The first item to cache.
	/// \param[in] lastItem The last item to cache.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::Raise_CacheItemsHint, GetItemInGroup
	virtual HRESULT STDMETHODCALLTYPE OnCacheHint(LVITEMINDEX firstItem, LVITEMINDEX lastItem) = 0;
};