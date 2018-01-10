//////////////////////////////////////////////////////////////////////
/// \class IItemComparator
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between the \c CompareItems callback method and the object that initiated the sorting</em>
///
/// This interface allows the \c CompareItems callback method to forward the message to the object
/// that initiated the sorting.
///
/// \sa ::CompareItems, ::CompareItemsEx, ExplorerListView::SortItems
//////////////////////////////////////////////////////////////////////


#pragma once


class IItemComparator
{
public:
	/// \brief <em>Compares two items by ID</em>
	///
	/// \param[in] itemID1 The unique ID of the first item to compare.
	/// \param[in] itemID2 The unique ID of the second item to compare.
	///
	/// \return -1 if the first item should precede the second; 0 if the items are equal; 1 if the
	///         second item should precede the first.
	///
	/// \sa CompareItemsEx, ::CompareItems, ExplorerListView::SortItems
	virtual int CompareItems(LONG itemID1, LONG itemID2) = 0;
	/// \brief <em>Compares two items by index</em>
	///
	/// \param[in] itemIndex1 The index of the first item to compare.
	/// \param[in] itemIndex2 The index of the second item to compare.
	///
	/// \return -1 if the first item should precede the second; 0 if the items are equal; 1 if the
	///         second item should precede the first.
	///
	/// \sa CompareItems, ::CompareItemsEx, ExplorerListView::SortItems
	virtual int CompareItemsEx(int itemIndex1, int itemIndex2) = 0;
};     // IItemComparator