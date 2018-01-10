//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView
//////////////////////////////////////////////////////////////////////


#pragma once

#include "ListViewColumn.h"
#include "ListViewColumns.h"
#include "ListViewFooterItem.h"
#include "ListViewFooterItems.h"
#include "ListViewGroup.h"
#include "ListViewGroups.h"
#include "ListViewItem.h"
#include "ListViewItems.h"
#include "ListViewSubItem.h"
#include "ListViewSubItems.h"
#include "ListViewWorkArea.h"
#include "ListViewWorkAreas.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates a \c ListViewColumn object</em>
	///
	/// Creates a \c ListViewColumn object that represents a given listview column and returns
	/// its \c IListViewColumn implementation as a smart pointer.
	///
	/// \param[in] columnIndex The index of the column for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewColumn object will work on.
	/// \param[in] validateColumnIndex If \c TRUE, the method will check whether the \c columnIndex
	///            parameter identifies an existing column; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewColumn implementation.
	static inline CComPtr<IListViewColumn> InitColumn(int columnIndex, ExplorerListView* pExLvw, BOOL validateColumnIndex = TRUE)
	{
		CComPtr<IListViewColumn> pColumn = NULL;
		InitColumn(columnIndex, pExLvw, IID_IListViewColumn, reinterpret_cast<IUnknown**>(&pColumn), validateColumnIndex);
		return pColumn;
	};

	/// \brief <em>Creates a \c ListViewColumn object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewColumn object that represents a given listview column and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] columnIndex The index of the column for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewColumn object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppColumn Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateColumnIndex If \c TRUE, the method will check whether the \c columnIndex
	///            parameter identifies an existing column; otherwise not.
	static inline void InitColumn(int columnIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppColumn, BOOL validateColumnIndex = TRUE)
	{
		ATLASSERT_POINTER(ppColumn, IUnknown*);
		ATLASSUME(ppColumn);

		*ppColumn = NULL;
		if(validateColumnIndex && !IsValidColumnIndex(columnIndex, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewColumn object and initialize it with the given parameters
		CComObject<ListViewColumn>* pLvwColumnObj = NULL;
		CComObject<ListViewColumn>::CreateInstance(&pLvwColumnObj);
		pLvwColumnObj->AddRef();
		pLvwColumnObj->SetOwner(pExLvw);
		pLvwColumnObj->Attach(columnIndex);
		pLvwColumnObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppColumn));
		pLvwColumnObj->Release();
	};

	/// \brief <em>Creates a \c ListViewColumns object</em>
	///
	/// Creates a \c ListViewColumns object that represents a collection of listview columns and returns
	/// its \c IListViewColumns implementation as a smart pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewColumns object will work on.
	///
	/// \return A smart pointer to the created object's \c IListViewColumns implementation.
	static inline CComPtr<IListViewColumns> InitColumns(ExplorerListView* pExLvw)
	{
		CComPtr<IListViewColumns> pColumns = NULL;
		InitColumns(pExLvw, IID_IListViewColumns, reinterpret_cast<IUnknown**>(&pColumns));
		return pColumns;
	};

	/// \brief <em>Creates a \c ListViewColumns object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewColumns object that represents a collection of listview columns and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewColumns object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppColumns Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitColumns(ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppColumns)
	{
		ATLASSERT_POINTER(ppColumns, IUnknown*);
		ATLASSUME(ppColumns);

		*ppColumns = NULL;
		// create a ListViewColumns object and initialize it with the given parameters
		CComObject<ListViewColumns>* pLvwColumnsObj = NULL;
		CComObject<ListViewColumns>::CreateInstance(&pLvwColumnsObj);
		pLvwColumnsObj->AddRef();
		pLvwColumnsObj->SetOwner(pExLvw);
		pLvwColumnsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppColumns));
		pLvwColumnsObj->Release();
	};

	/// \brief <em>Creates a \c ListViewFooterItem object</em>
	///
	/// Creates a \c ListViewFooterItem object that represents a given listview footer item and returns
	/// its \c IListViewFooterItem implementation as a smart pointer.
	///
	/// \param[in] itemIndex The zero-based index of the footer item for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewFooterItem object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex
	///            parameter identifies an existing footer item; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewFooterItem implementation.
	static inline CComPtr<IListViewFooterItem> InitFooterItem(int itemIndex, ExplorerListView* pExLvw, BOOL validateItemIndex = TRUE)
	{
		CComPtr<IListViewFooterItem> pFooterItem = NULL;
		InitFooterItem(itemIndex, pExLvw, IID_IListViewFooterItem, reinterpret_cast<IUnknown**>(&pFooterItem), validateItemIndex);
		return pFooterItem;
	};

	/// \brief <em>Creates a \c ListViewFooterItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewFooterItem object that represents a given listview footer item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemIndex The zero-based index of the footer item for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewFooterItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppFooterItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex
	///            parameter identifies an existing footer item; otherwise not.
	static inline void InitFooterItem(int itemIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppFooterItem, BOOL validateItemIndex = TRUE)
	{
		ATLASSERT_POINTER(ppFooterItem, IUnknown*);
		ATLASSUME(ppFooterItem);

		*ppFooterItem = NULL;
		if(validateItemIndex && !IsValidFooterItemIndex(itemIndex, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewFooterItem object and initialize it with the given parameters
		CComObject<ListViewFooterItem>* pLvwFooterItemObj = NULL;
		CComObject<ListViewFooterItem>::CreateInstance(&pLvwFooterItemObj);
		pLvwFooterItemObj->AddRef();
		pLvwFooterItemObj->SetOwner(pExLvw);
		pLvwFooterItemObj->Attach(itemIndex);
		pLvwFooterItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppFooterItem));
		pLvwFooterItemObj->Release();
	};

	/// \brief <em>Creates a \c ListViewFooterItems object</em>
	///
	/// Creates a \c ListViewFooterItems object that represents a collection of listview footer items and
	/// returns its \c IListViewFooterItems implementation as a smart pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewFooterItems object will work on.
	///
	/// \return A smart pointer to the created object's \c IListViewFooterItems implementation.
	static inline CComPtr<IListViewFooterItems> InitFooterItems(ExplorerListView* pExLvw)
	{
		CComPtr<IListViewFooterItems> pFooterItems = NULL;
		InitFooterItems(pExLvw, IID_IListViewFooterItems, reinterpret_cast<IUnknown**>(&pFooterItems));
		return pFooterItems;
	};

	/// \brief <em>Creates a \c ListViewFooterItems object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewFooterItems object that represents a collection of listview footer items and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewFooterItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppFooterItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitFooterItems(ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppFooterItems)
	{
		ATLASSERT_POINTER(ppFooterItems, IUnknown*);
		ATLASSUME(ppFooterItems);

		*ppFooterItems = NULL;
		// create a ListViewFooterItems object and initialize it with the given parameters
		CComObject<ListViewFooterItems>* pLvwFooterItemsObj = NULL;
		CComObject<ListViewFooterItems>::CreateInstance(&pLvwFooterItemsObj);
		pLvwFooterItemsObj->AddRef();
		pLvwFooterItemsObj->SetOwner(pExLvw);
		pLvwFooterItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppFooterItems));
		pLvwFooterItemsObj->Release();
	};

	/// \brief <em>Creates a \c ListViewGroup object</em>
	///
	/// Creates a \c ListViewGroup object that represents a given listview group and returns
	/// its \c IListViewGroup implementation as a smart pointer.
	///
	/// \param[in] groupID The unique ID of the group for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewGroup object will work on.
	/// \param[in] validateGroupID If \c TRUE, the method will check whether the \c groupID
	///            parameter identifies an existing group; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewGroup implementation.
	///
	/// \sa InitGroupByIndex
	static inline CComPtr<IListViewGroup> InitGroup(int groupID, ExplorerListView* pExLvw, BOOL validateGroupID = TRUE)
	{
		CComPtr<IListViewGroup> pGroup = NULL;
		InitGroup(groupID, pExLvw, IID_IListViewGroup, reinterpret_cast<IUnknown**>(&pGroup), validateGroupID);
		return pGroup;
	};

	/// \brief <em>Creates a \c ListViewGroup object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewGroup object that represents a given listview group and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] groupID The unique ID of the group for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewGroup object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppGroup Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateGroupID If \c TRUE, the method will check whether the \c groupID
	///            parameter identifies an existing group; otherwise not.
	///
	/// \sa InitGroupByIndex
	static inline void InitGroup(int groupID, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppGroup, BOOL validateGroupID = TRUE)
	{
		ATLASSERT_POINTER(ppGroup, IUnknown*);
		ATLASSUME(ppGroup);

		*ppGroup = NULL;
		if(validateGroupID && !IsValidGroupID(groupID, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewGroup object and initialize it with the given parameters
		CComObject<ListViewGroup>* pLvwGroupObj = NULL;
		CComObject<ListViewGroup>::CreateInstance(&pLvwGroupObj);
		pLvwGroupObj->AddRef();
		pLvwGroupObj->SetOwner(pExLvw);
		pLvwGroupObj->Attach(groupID);
		pLvwGroupObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppGroup));
		pLvwGroupObj->Release();
	};

	/// \brief <em>Creates a \c ListViewGroup object</em>
	///
	/// Creates a \c ListViewGroup object that represents a given listview group and returns
	/// its \c IListViewGroup implementation as a smart pointer.
	///
	/// \param[in] groupIndex The index of the group for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewGroup object will work on.
	///
	/// \return A smart pointer to the created object's \c IListViewGroup implementation.
	///
	/// \sa InitGroup
	static inline CComPtr<IListViewGroup> InitGroupByIndex(int groupIndex, ExplorerListView* pExLvw)
	{
		CComPtr<IListViewGroup> pGroup = NULL;
		InitGroupByIndex(groupIndex, pExLvw, IID_IListViewGroup, reinterpret_cast<IUnknown**>(&pGroup));
		return pGroup;
	};

	/// \brief <em>Creates a \c ListViewGroup object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewGroup object that represents a given listview group and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] groupIndex The index of the group for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewGroup object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppGroup Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa InitGroup
	static inline void InitGroupByIndex(int groupIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppGroup)
	{
		ATLASSERT_POINTER(ppGroup, IUnknown*);
		ATLASSUME(ppGroup);

		HWND hWndLvw = NULL;
		if(pExLvw) {
			OLE_HANDLE h = NULL;
			pExLvw->get_hWnd(&h);
			hWndLvw = static_cast<HWND>(LongToHandle(h));
		}

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		group.mask = LVGF_GROUPID;
		if(hWndLvw && SendMessage(hWndLvw, LVM_GETGROUPINFOBYINDEX, groupIndex, reinterpret_cast<LPARAM>(&group))) {
			InitGroup(group.iGroupId, pExLvw, requiredInterface, ppGroup);
		} else {
			*ppGroup = NULL;
		}
	};

	/// \brief <em>Creates a \c ListViewGroups object</em>
	///
	/// Creates a \c ListViewGroups object that represents a collection of listview groups and returns
	/// its \c IListViewGroups implementation as a smart pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewGroups object will work on.
	///
	/// \return A smart pointer to the created object's \c IListViewGroups implementation.
	static inline CComPtr<IListViewGroups> InitGroups(ExplorerListView* pExLvw)
	{
		CComPtr<IListViewGroups> pGroups = NULL;
		InitGroups(pExLvw, IID_IListViewGroups, reinterpret_cast<IUnknown**>(&pGroups));
		return pGroups;
	};

	/// \brief <em>Creates a \c ListViewGroups object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewGroups object that represents a collection of listview groups and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewGroups object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppGroups Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitGroups(ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppGroups)
	{
		ATLASSERT_POINTER(ppGroups, IUnknown*);
		ATLASSUME(ppGroups);

		*ppGroups = NULL;
		// create a ListViewGroups object and initialize it with the given parameters
		CComObject<ListViewGroups>* pLvwGroupsObj = NULL;
		CComObject<ListViewGroups>::CreateInstance(&pLvwGroupsObj);
		pLvwGroupsObj->AddRef();
		pLvwGroupsObj->SetOwner(pExLvw);
		pLvwGroupsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppGroups));
		pLvwGroupsObj->Release();
	};

	/// \brief <em>Creates a \c ListViewItem object</em>
	///
	/// Creates a \c ListViewItem object that represents a given listview item and returns
	/// its \c IListViewItem implementation as a smart pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewItem object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewItem implementation.
	static inline CComPtr<IListViewItem> InitListItem(LVITEMINDEX& itemIndex, ExplorerListView* pExLvw, BOOL validateItemIndex = TRUE)
	{
		CComPtr<IListViewItem> pItem = NULL;
		InitListItem(itemIndex, pExLvw, IID_IListViewItem, reinterpret_cast<IUnknown**>(&pItem), validateItemIndex);
		return pItem;
	};

	/// \brief <em>Creates a \c ListViewItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewItem object that represents a given listview item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] itemIndex The index of the item for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c itemIndex parameter
	///            identifies an existing item; otherwise not.
	static inline void InitListItem(LVITEMINDEX& itemIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppItem, BOOL validateItemIndex = TRUE)
	{
		ATLASSERT_POINTER(ppItem, IUnknown*);
		ATLASSUME(ppItem);

		*ppItem = NULL;
		if(validateItemIndex && !IsValidItemIndex(itemIndex.iItem, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewItem object and initialize it with the given parameters
		CComObject<ListViewItem>* pLvwItemObj = NULL;
		CComObject<ListViewItem>::CreateInstance(&pLvwItemObj);
		pLvwItemObj->AddRef();
		pLvwItemObj->SetOwner(pExLvw);
		pLvwItemObj->Attach(itemIndex);
		pLvwItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItem));
		pLvwItemObj->Release();
	};

	/// \brief <em>Creates a \c ListViewItems object</em>
	///
	/// Creates a \c ListViewItems object that represents a collection of listview items and returns
	/// its \c IListViewItems implementation as a smart pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewItems object will work on.
	///
	/// \return A smart pointer to the created object's \c IListViewItems implementation.
	static inline CComPtr<IListViewItems> InitListItems(ExplorerListView* pExLvw)
	{
		CComPtr<IListViewItems> pItems = NULL;
		InitListItems(pExLvw, IID_IListViewItems, reinterpret_cast<IUnknown**>(&pItems));
		return pItems;
	};

	/// \brief <em>Creates a \c ListViewItems object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewItems object that represents a collection of listview items and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitListItems(ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppItems)
	{
		ATLASSERT_POINTER(ppItems, IUnknown*);
		ATLASSUME(ppItems);

		*ppItems = NULL;
		// create a ListViewItems object and initialize it with the given parameters
		CComObject<ListViewItems>* pLvwItemsObj = NULL;
		CComObject<ListViewItems>::CreateInstance(&pLvwItemsObj);
		pLvwItemsObj->AddRef();

		pLvwItemsObj->SetOwner(pExLvw);

		pLvwItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppItems));
		pLvwItemsObj->Release();
	};

	/// \brief <em>Creates a \c ListViewSubItem object</em>
	///
	/// Creates a \c ListViewSubItem object that represents a given sub-item and returns
	/// its \c IListViewSubItem implementation as a smart pointer.
	///
	/// \param[in] parentItemIndex The parent item of the sub-item for which to create the object.
	/// \param[in] subItemIndex The index of the sub-item for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewSubItem object will work on.
	/// \param[in] validateParentItemIndex If \c TRUE, the method will check whether the \c parentItemIndex
	///            parameter identifies an existing item; otherwise not.
	/// \param[in] validateSubItemIndex If \c TRUE, the method will check whether the \c subItemIndex
	///            parameter identifies an existing sub-item; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewSubItem implementation.
	static inline CComPtr<IListViewSubItem> InitListSubItem(LVITEMINDEX& parentItemIndex, int subItemIndex, ExplorerListView* pExLvw, BOOL validateParentItemIndex = TRUE, BOOL validateSubItemIndex = TRUE)
	{
		CComPtr<IListViewSubItem> pSubItem = NULL;
		InitListSubItem(parentItemIndex, subItemIndex, pExLvw, IID_IListViewSubItem, reinterpret_cast<IUnknown**>(&pSubItem), validateParentItemIndex, validateSubItemIndex);
		return pSubItem;
	};

	/// \brief <em>Creates a \c ListViewSubItem object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewSubItem object that represents a given sub-item and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] parentItemIndex The parent item of the sub-item for which to create the object.
	/// \param[in] subItemIndex The index of the sub-item for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewSubItem object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppSubItem Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateParentItemIndex If \c TRUE, the method will check whether the \c parentItemIndex
	///            parameter identifies an existing item; otherwise not.
	/// \param[in] validateSubItemIndex If \c TRUE, the method will check whether the \c subItemIndex
	///            parameter identifies an existing sub-item; otherwise not.
	static inline void InitListSubItem(LVITEMINDEX& parentItemIndex, int subItemIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppSubItem, BOOL validateParentItemIndex = TRUE, BOOL validateSubItemIndex = TRUE)
	{
		ATLASSERT_POINTER(ppSubItem, IUnknown*);
		ATLASSUME(ppSubItem);

		*ppSubItem = NULL;
		if(validateParentItemIndex && !IsValidItemIndex(parentItemIndex.iItem, pExLvw)) {
			// there's nothing useful we could return
			return;
		}
		if(validateSubItemIndex && !IsValidColumnIndex(subItemIndex, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewSubItem object and initialize it with the given parameters
		CComObject<ListViewSubItem>* pLvwSubItemObj = NULL;
		CComObject<ListViewSubItem>::CreateInstance(&pLvwSubItemObj);
		pLvwSubItemObj->AddRef();
		pLvwSubItemObj->SetOwner(pExLvw);
		pLvwSubItemObj->Attach(parentItemIndex, subItemIndex);
		pLvwSubItemObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppSubItem));
		pLvwSubItemObj->Release();
	};

	/// \brief <em>Creates a \c ListViewSubItems object</em>
	///
	/// Creates a \c ListViewSubItems object that represents a collection of listview sub-items and returns
	/// its \c IListViewSubItems implementation as a smart pointer.
	///
	/// \param[in] parentItemIndex The parent item of the sub-items for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewSubItems object will work on.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c parentItemIndex
	///            parameter identifies an existing item; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewSubItems implementation.
	static inline CComPtr<IListViewSubItems> InitListSubItems(LVITEMINDEX& parentItemIndex, ExplorerListView* pExLvw, BOOL validateItemIndex = FALSE)
	{
		CComPtr<IListViewSubItems> pSubItems = NULL;
		InitListSubItems(parentItemIndex, pExLvw, IID_IListViewSubItems, reinterpret_cast<IUnknown**>(&pSubItems), validateItemIndex);
		return pSubItems;
	};

	/// \brief <em>Creates a \c ListViewSubItems object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewSubItems object that represents a collection of listview sub-items and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] parentItemIndex The parent item of the sub-items for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewSubItems object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppSubItems Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateItemIndex If \c TRUE, the method will check whether the \c parentItemIndex
	///            parameter identifies an existing item; otherwise not.
	static inline void InitListSubItems(LVITEMINDEX& parentItemIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppSubItems, BOOL validateItemIndex = FALSE)
	{
		ATLASSERT_POINTER(ppSubItems, IUnknown*);
		ATLASSUME(ppSubItems);

		*ppSubItems = NULL;
		if(validateItemIndex && !IsValidItemIndex(parentItemIndex.iItem, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewSubItems object and initialize it with the given parameters
		CComObject<ListViewSubItems>* pLvwSubItemsObj = NULL;
		CComObject<ListViewSubItems>::CreateInstance(&pLvwSubItemsObj);
		pLvwSubItemsObj->AddRef();

		pLvwSubItemsObj->SetOwner(pExLvw);
		pLvwSubItemsObj->properties.parentItemIndex = parentItemIndex;

		pLvwSubItemsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppSubItems));
		pLvwSubItemsObj->Release();
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static inline CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<IUnknown**>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static inline void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, IUnknown** ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, IUnknown*);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};

	/// \brief <em>Creates a \c ListViewWorkArea object</em>
	///
	/// Creates a \c ListViewWorkArea object that represents a given listview working area and returns
	/// its \c IListViewWorkArea implementation as a smart pointer.
	///
	/// \param[in] workAreaIndex The index of the working area for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewWorkArea object will work on.
	/// \param[in] validateWorkAreaIndex If \c TRUE, the method will check whether the \c workAreaIndex
	///            parameter identifies an existing working area; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IListViewWorkArea implementation.
	static inline CComPtr<IListViewWorkArea> InitWorkArea(int workAreaIndex, ExplorerListView* pExLvw, BOOL validateWorkAreaIndex = TRUE)
	{
		CComPtr<IListViewWorkArea> pWorkArea = NULL;
		InitWorkArea(workAreaIndex, pExLvw, IID_IListViewWorkArea, reinterpret_cast<IUnknown**>(&pWorkArea), validateWorkAreaIndex);
		return pWorkArea;
	};

	/// \brief <em>Creates a \c ListViewWorkArea object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewWorkArea object that represents a given listview working area and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] workAreaIndex The index of the working area for which to create the object.
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewWorkArea object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppWorkArea Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateWorkAreaIndex If \c TRUE, the method will check whether the \c workAreaIndex
	///            parameter identifies an existing working area; otherwise not.
	static inline void InitWorkArea(int workAreaIndex, ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppWorkArea, BOOL validateWorkAreaIndex = TRUE)
	{
		ATLASSERT_POINTER(ppWorkArea, IUnknown*);
		ATLASSUME(ppWorkArea);

		*ppWorkArea = NULL;
		if(validateWorkAreaIndex && !IsValidWorkAreaIndex(workAreaIndex, pExLvw)) {
			// there's nothing useful we could return
			return;
		}

		// create a ListViewWorkArea object and initialize it with the given parameters
		CComObject<ListViewWorkArea>* pLvwWorkAreaObj = NULL;
		CComObject<ListViewWorkArea>::CreateInstance(&pLvwWorkAreaObj);
		pLvwWorkAreaObj->AddRef();
		pLvwWorkAreaObj->SetOwner(pExLvw);
		pLvwWorkAreaObj->Attach(workAreaIndex);
		pLvwWorkAreaObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppWorkArea));
		pLvwWorkAreaObj->Release();
	};

	/// \brief <em>Creates a \c ListViewWorkAreas object</em>
	///
	/// Creates a \c ListViewWorkAreas object that represents a collection of listview working areas and
	/// returns its \c IListViewWorkAreas implementation as a smart pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewWorkAreas object will work on.
	///
	/// \return A smart pointer to the created object's \c IListViewWorkAreas implementation.
	static inline CComPtr<IListViewWorkAreas> InitWorkAreas(ExplorerListView* pExLvw)
	{
		CComPtr<IListViewWorkAreas> pWorkAreas = NULL;
		InitWorkAreas(pExLvw, IID_IListViewWorkAreas, reinterpret_cast<IUnknown**>(&pWorkAreas));
		return pWorkAreas;
	};

	/// \brief <em>Creates a \c ListViewWorkAreas object</em>
	///
	/// \overload
	///
	/// Creates a \c ListViewWorkAreas object that represents a collection of listview working areas and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pExLvw The \c ExplorerListView object the \c ListViewWorkAreas object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppWorkAreas Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitWorkAreas(ExplorerListView* pExLvw, REFIID requiredInterface, IUnknown** ppWorkAreas)
	{
		ATLASSERT_POINTER(ppWorkAreas, IUnknown*);
		ATLASSUME(ppWorkAreas);

		*ppWorkAreas = NULL;
		// create a ListViewWorkAreas object and initialize it with the given parameters
		CComObject<ListViewWorkAreas>* pLvwWorkAreasObj = NULL;
		CComObject<ListViewWorkAreas>::CreateInstance(&pLvwWorkAreasObj);
		pLvwWorkAreasObj->AddRef();
		pLvwWorkAreasObj->SetOwner(pExLvw);
		pLvwWorkAreasObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppWorkAreas));
		pLvwWorkAreasObj->Release();
	};
};     // ClassFactory