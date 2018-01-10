//////////////////////////////////////////////////////////////////////
/// \class IItemContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between a \c ListViewItemContainer object and its creator object</em>
///
/// This interface allows an \c ExplorerListView object to inform a \c ListViewItemContainer object
/// about item deletions.
///
/// \sa ExplorerListView::RegisterItemContainer, ListViewItemContainer
//////////////////////////////////////////////////////////////////////


#pragma once


class IItemContainer
{
public:
	/// \brief <em>Adds a large number of items to the container</em>
	///
	/// \param[in] numberOfItems The number of items to add.
	/// \param[in] pItemIDs An array containing the unique IDs of the items to add.
	///
	/// \remarks This method does not check whether any of the specified items already is in the container.
	///
	/// \sa RemovedItem, ExplorerListView::OnInternalCreateItemContainer
	virtual void AddItems(UINT numberOfItems, __in_ecount(numberOfItems) PLONG pItemIDs) = 0;
	/// \brief <em>An item was removed and the item container should check its content</em>
	///
	/// \param[in] itemID The unique ID of the removed item. If 0, all items were removed.
	///
	/// \sa AddItems, ExplorerListView::RemoveItemFromItemContainers
	virtual void RemovedItem(LONG itemID) = 0;
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa ExplorerListView::DeregisterItemContainer, containerID
	virtual DWORD GetID(void) = 0;

	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
};     // IItemContainer