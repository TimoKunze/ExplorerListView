//////////////////////////////////////////////////////////////////////
/// \class ListViewFooterItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListViewFooterItem objects</em>
///
/// This class provides easy access to collections of \c ListViewFooterItem objects. A
/// \c ListViewFooterItems object is used to group the control's footer items.
///
/// \remarks Requires comctl32.dll version 6.10 or higher.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewFooterItems, ListViewFooterItem, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewFooterItems, ListViewFooterItem, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IListViewFooterItemsEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "ListViewFooterItem.h"


class ATL_NO_VTABLE ListViewFooterItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewFooterItems, &CLSID_ListViewFooterItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewFooterItems>,
    public Proxy_IListViewFooterItemsEvents<ListViewFooterItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListViewFooterItems, &IID_IListViewFooterItems, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewFooterItems, &IID_IListViewFooterItems, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWFOOTERITEMS)

		BEGIN_COM_MAP(ListViewFooterItems)
			COM_INTERFACE_ENTRY(IListViewFooterItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewFooterItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewFooterItemsEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the footer items</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ListViewFooterItem objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x footer items</em>
	///
	/// Retrieves the next \c numberOfMaxItems footer items from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of footer items the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a footer item's \c IListViewFooterItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of footer items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListViewFooterItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// footer item in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x footer items</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip footer items.
	///
	/// \param[in] numberOfItemsToSkip The number of footer items to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IListViewFooterItems
	///
	//@{
	/// \brief <em>Retrieves a \c ListViewFooterItem object from the collection</em>
	///
	/// Retrieves a \c ListViewFooterItem object from the collection that wraps the listview footer item
	/// identified by \c footerItemIdentifier.
	///
	/// \param[in] footerItemIdentifier A value that identifies the listview footer item to be retrieved.
	/// \param[in] footerItemIdentifierType A value specifying the meaning of \c footerItemIdentifier. Any of
	///            the values defined by the \c FooterItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppFooterItem Receives the footer item's \c IListViewFooterItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListViewFooterItem, Add, ExLVwLibU::FooterItemIdentifierTypeConstants
	/// \else
	///   \sa ListViewFooterItem, Add, ExLVwLibA::FooterItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG footerItemIdentifier, FooterItemIdentifierTypeConstants footerItemIdentifierType = fiitIndex, IListViewFooterItem** ppFooterItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListViewFooterItem objects
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

	/// \brief <em>Adds a footer item to the listview</em>
	///
	/// Adds a footer item with the specified properties at the specified position in the control and returns
	/// a \c ListViewFooterItem object wrapping the inserted footer item.
	///
	/// \param[in] footerItemText The new footer item's text. The maximum number of characters in this text
	///            is \c MAX_FOOTERITEMTEXTLENGTH. With current versions of comctl32.dll a footer item's text
	///            cannot be changed after the item has been inserted.
	/// \param[in] insertAt The new footer item's zero-based index. If set to -1, the footer item will be
	///            inserted as the last footer item.
	/// \param[in] iconIndex The zero-based index of the footer item's icon in the control's
	///            \c ilFooterItems imagelist. With current versions of comctl32.dll a footer item's
	///            icon index cannot be retrieved or changed after the item has been inserted.
	/// \param[in] itemData A \c LONG value that will be associated with the footer item. With current
	///            versions of comctl32.dll a footer item's associated data cannot be changed after the
	///            item has been inserted. Also the events \c FooterItemClick and \c FreeFooterItemData are
	///            not fired for footer items whose associated data is 0.
	/// \param[out] ppAddedFooterItem Receives the added footer item's \c IListViewFooterItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks With current versions of comctl32.dll it is not possible to add more than 4 footer items.
	///
	/// \if UNICODE
	///   \sa Count, RemoveAll, ListViewFooterItem::get_Text, ListViewFooterItem::get_Index,
	///       ListViewFooterItem::get_ItemData, ExplorerListView::get_hImageList,
	///       MAX_FOOTERITEMTEXTLENGTH, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa Count, RemoveAll, ListViewFooterItem::get_Text, ListViewFooterItem::get_Index,
	///       ListViewFooterItem::get_ItemData, ExplorerListView::get_hImageList,
	///       MAX_FOOTERITEMTEXTLENGTH, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR footerItemText, LONG insertAt = -1, LONG iconIndex = 0, LONG itemData = 1, IListViewFooterItem** ppAddedFooterItem = NULL);
	/// \brief <em>Counts the footer items in the collection</em>
	///
	/// Retrieves the number of \c ListViewFooterItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes all footer items in the collection from the listview</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This also hides the footer area's caption.
	///
	/// \sa Add, Count
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

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
		/// \brief <em>Holds the position index of the next enumerated footer item</em>
		int nextEnumeratedFooterItem;

		Properties()
		{
			pOwnerExLvw = NULL;
			nextEnumeratedFooterItem = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains the footer items in this collection.
		///
		/// \sa pOwnerExLvw, WindowQueryInterface
		HWND GetExLvwHWnd(void);
		/// \brief <em>Queries the owning listview window for the specified interface</em>
		///
		/// \param[in] requiredInterface The IID of the interface of which the control's implementation will
		///            be returned.
		/// \param[out] ppObject Receives the control's implementation of the interface identified by
		///             \c requiredInterface.
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa GetExLvwHWnd
		HRESULT WindowQueryInterface(REFIID requiredInterface, __deref_out LPUNKNOWN* ppObject);
	} properties;
};     // ListViewFooterItems

OBJECT_ENTRY_AUTO(__uuidof(ListViewFooterItems), ListViewFooterItems)