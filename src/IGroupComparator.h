//////////////////////////////////////////////////////////////////////
/// \class IGroupComparator
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between the \c CompareGroups callback method and the object that initiated the sorting</em>
///
/// This interface allows the \c CompareGroups callback method to forward the message to the object
/// that initiated the sorting.
///
/// \sa ::CompareGroups, ExplorerListView::SortGroups
//////////////////////////////////////////////////////////////////////


#pragma once


class IGroupComparator
{
public:
	/// \brief <em>Compares two groups by ID</em>
	///
	/// \param[in] groupID1 The unique ID of the first group to compare.
	/// \param[in] groupID2 The unique ID of the second group to compare.
	///
	/// \return -1 if the first group should precede the second; 0 if the groups are equal; 1 if the
	///         second group should precede the first.
	///
	/// \sa ::CompareGroups, ExplorerListView::SortGroups
	virtual int CompareGroups(int groupID1, int groupID2) = 0;
};     // IGroupComparator