//////////////////////////////////////////////////////////////////////
/// \class ListViewSubItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListViewSubItem objects</em>
///
/// This class provides easy access to collections of \c ListViewSubItem objects. A \c ListViewSubItems
/// object is used to group a listview item's sub-items.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewSubItems, ListViewSubItem, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewSubItems, ListViewSubItem, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IListViewSubItemsEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "ListViewSubItem.h"


class ATL_NO_VTABLE ListViewSubItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewSubItems, &CLSID_ListViewSubItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewSubItems>,
    public Proxy_IListViewSubItemsEvents<ListViewSubItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListViewSubItems, &IID_IListViewSubItems, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewSubItems, &IID_IListViewSubItems, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWSUBITEMS)

		BEGIN_COM_MAP(ListViewSubItems)
			COM_INTERFACE_ENTRY(IListViewSubItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewSubItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewSubItemsEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the items</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ListViewSubItem objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x items</em>
	///
	/// Retrieves the next \c numberOfMaxItems items from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of items the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a item's \c IListViewSubItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListViewSubItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// item in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x items</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip items.
	///
	/// \param[in] numberOfItemsToSkip The number of items to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IListViewSubItems
	///
	//@{
	/// \brief <em>Retrieves a \c ListViewSubItem object from the collection</em>
	///
	/// Retrieves a \c ListViewSubItem object from the collection that wraps the sub-item identified
	/// by \c subItemIdentifier.
	///
	/// \param[in] subItemIdentifier A value that identifies the sub-item to be retrieved.
	/// \param[in] subItemIdentifierType A value specifying the meaning of \c subItemIdentifier. Any of
	///            the values defined by the \c ColumnIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppSubItem Receives the sub-item's \c IListViewSubItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListViewSubItem, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa ListViewSubItem, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG subItemIdentifier, ColumnIdentifierTypeConstants subItemIdentifierType = citIndex, IListViewSubItem** ppSubItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListViewSubItem objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);
	/// \brief <em>Retrieves all sub-items' parent item</em>
	///
	/// Retrieves a \c ListViewItem object wrapping the listview item that is the parent item
	/// of all sub-items in the collection.
	///
	/// \param[out] ppParentItem Receives the parent item's \c IListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ListViewItem
	virtual HRESULT STDMETHODCALLTYPE get_ParentItem(IListViewItem** ppParentItem);

	/// \brief <em>Counts the sub-items in the collection</em>
	///
	/// Retrieves the number of \c ListViewSubItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given item's sub-items</em>
	///
	/// Attaches this object to a given item's sub-items, so that the sub-items can be managed using
	/// this object's methods.
	///
	/// \param[in] parentItemIndex The parent item of the sub-items to attach to.
	///
	/// \sa Detach
	void Attach(LVITEMINDEX& parentItemIndex);
	/// \brief <em>Detaches this object from an item's sub-items</em>
	///
	/// Detaches this object from the item's sub-items it currently wraps, so that it doesn't wrap any
	/// sub-items anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this collection</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The index of the item the sub-items wrapped by this object belong to</em>
		LVITEMINDEX parentItemIndex;
		/// \brief <em>Holds the next enumerated sub-item</em>
		int nextEnumeratedSubItem;

		Properties()
		{
			pOwnerExLvw = NULL;
			parentItemIndex.iItem = -1;
			parentItemIndex.iGroup = 0;
			nextEnumeratedSubItem = 1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains the sub-items in this collection.
		///
		/// \sa pOwnerExLvw, GetHeaderHWnd
		HWND GetExLvwHWnd(void);
		/// \brief <em>Retrieves the window handle of the owning listview's header control</em>
		///
		/// \return The window handle of the header control containing the columns that the sub-items in
		///         this collection refer to.
		///
		/// \sa pOwnerExLvw, GetExLvwHWnd
		HWND GetHeaderHWnd(void);
	} properties;
};     // ListViewSubItems

OBJECT_ENTRY_AUTO(__uuidof(ListViewSubItems), ListViewSubItems)