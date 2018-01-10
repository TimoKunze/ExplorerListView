//////////////////////////////////////////////////////////////////////
/// \class ListViewItems
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListViewItem objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ListViewItem
/// objects. A \c ListViewItems object is used to group items that have certain properties in
/// common.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewItems, ListViewItem, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewItems, ListViewItem, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IListViewItemsEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "ListViewItem.h"


class ATL_NO_VTABLE ListViewItems : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewItems, &CLSID_ListViewItems>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewItems>,
    public Proxy_IListViewItemsEvents<ListViewItems>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListViewItems, &IID_IListViewItems, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewItems, &IID_IListViewItems, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ClassFactory;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ListViewItems();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWITEMS)

		BEGIN_COM_MAP(ListViewItems)
			COM_INTERFACE_ENTRY(IListViewItems)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewItems)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewItemsEvents))
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
	/// the \c ListViewItem objects managed by this collection object.
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
	///                the pointer to a item's \c IListViewItem implementation.
	/// \param[out] pNumberOfItemsReturned The number of items that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListViewItem,
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
	/// \name Implementation of IListViewItems
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CaseSensitiveFilters property</em>
	///
	/// Retrieves whether string comparisons, that are done when applying the filters on an item, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CaseSensitiveFilters, get_Filter, get_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE get_CaseSensitiveFilters(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CaseSensitiveFilters property</em>
	///
	/// Sets whether string comparisons, that are done when applying the filters on an item, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaseSensitiveFilters, put_Filter, put_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE put_CaseSensitiveFilters(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ComparisonFunction property</em>
	///
	/// Retrieves an item filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T itemProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG, a \c SAFEARRAY of \c LONG,
	/// \c BSTR, \c IListViewGroup or \c IListViewWorkArea). This function must compare its arguments and
	/// return a non-zero value if the arguments are equal and zero otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters, ListViewGroup, ListViewWorkArea,
	///       ExLVwLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters, ListViewGroup, ListViewWorkArea,
	///       ExLVwLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue);
	/// \brief <em>Sets the \c ComparisonFunction property</em>
	///
	/// Sets an item filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T itemProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG, a \c SAFEARRAY of \c LONG,
	/// \c BSTR, \c IListViewGroup or \c IListViewWorkArea). This function must compare its arguments and
	/// return a non-zero value if the arguments are equal and zero otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters, ListViewGroup, ListViewWorkArea,
	///       ExLVwLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters, ListViewGroup, ListViewWorkArea,
	///       ExLVwLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves an item filter.\n
	/// An \c IListViewItems collection can be filtered by any of \c IListViewItem's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, ExLVwLibU::FilteredPropertyConstants
	/// \else
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, ExLVwLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets an item filter.\n
	/// An \c IListViewItems collection can be filtered by any of \c IListViewItem's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       ExLVwLibU::FilteredPropertyConstants
	/// \else
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       ExLVwLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c FilterType property</em>
	///
	/// Retrieves an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_FilterType, get_Filter, ExLVwLibU::FilteredPropertyConstants,
	///       ExLVwLibU::FilterTypeConstants
	/// \else
	///   \sa put_FilterType, get_Filter, ExLVwLibA::FilteredPropertyConstants,
	///       ExLVwLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue);
	/// \brief <em>Sets the \c FilterType property</em>
	///
	/// Sets an item filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, ExLVwLibU::FilteredPropertyConstants,
	///       ExLVwLibU::FilterTypeConstants
	/// \else
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, ExLVwLibA::FilteredPropertyConstants,
	///       ExLVwLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue);
	/// \brief <em>Retrieves a \c ListViewItem object from the collection</em>
	///
	/// Retrieves a \c ListViewItem object from the collection that wraps the listview item identified
	/// by \c itemIdentifier.
	///
	/// \param[in] itemIdentifier A value that identifies the listview item to be retrieved.
	/// \param[in] groupIndex If the control is in virtual mode and \c itemIdentifier specifies the item's
	///            zero-based index, this parameter specifies the zero-based index of the group containing
	///            the item.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppItem Receives the item's \c IListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If \c itemIdentifierType is set to \c iitIndex and \c itemIdentifier is set to -1 and the
	///          collection doesn't use filters, the returned object may be used to control some specific
	///          properties of all items at once.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListViewItem::get_GroupIndex, ExplorerListView::get_VirtualMode, Add, Remove, Contains,
	///       ExLVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa ListViewItem::get_GroupIndex, ExplorerListView::get_VirtualMode, Add, Remove, Contains,
	///       ExLVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG itemIdentifier, LONG groupIndex = 0, ItemIdentifierTypeConstants itemIdentifierType = iitIndex, IListViewItem** ppItem = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListViewItem objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds an item to the listview</em>
	///
	/// Adds an item with the specified properties at the specified position in the control and returns a
	/// \c ListViewItem object wrapping the inserted item.
	///
	/// \param[in] itemText The new item's caption text. The maximum number of characters in this text
	///            is \c MAX_ITEMTEXTLENGTH. If set to \c NULL, the control will fire the
	///            \c ItemGetDisplayInfo event each time this property's value is required.
	/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted
	///            as the last item.
	/// \param[in] iconIndex The zero-based index of the item's icon in the control's \c ilSmall, \c ilLarge,
	///            \c ilExtraLarge and \c ilHighResolution imagelists. If set to -1, the control will fire
	///            the \c ItemGetDisplayInfo event each time this property's value is required. A value of -2
	///            means 'not specified' and is valid if there's no imagelist associated with the control.
	/// \param[in] itemIndentation The new item's indentation in 'Details' view in image widths. If set
	///            to 1, the item's indentation will be 1 image width; if set to 2, it will be 2 image
	///            widths and so on. If set to -1, the control will fire the \c ItemGetDisplayInfo event
	///            each time this property's value is required.
	/// \param[in] itemData A \c LONG value that will be associated with the item.
	/// \param[in] groupID The unique ID of the group that the new item will belong to. If set to \c -2,
	///            the item won't belong to any group. With comctl32.dll version 6.10 or higher, if set to
	///            -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value
	///            is required. However, this feature seems to be broken in comctl32.dll, so that the event
	///            is raised only once per item.
	/// \param[in] tileViewColumns An array of \c TILEVIEWSUBITEM structs which define the sub-items that
	///            will be displayed below the new item's text in 'Tiles' and 'Extended Tiles' view. If set
	///            to an empty array, no details will be displayed. If set to \c Empty, the control will
	///            fire the \c ItemGetDisplayInfo event each time this property's value is required.
	/// \param[out] ppAddedItem Receives the added item's \c IListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c groupID and \c tileViewColumns parameters will be ignored if comctl32.dll is
	///          used in a version older than 6.0.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, ListViewItem::get_Text, ListViewItem::get_IconIndex,
	///       ListViewItem::get_Indent, ListViewItem::get_ItemData, ListViewItem::get_Group,
	///       ListViewGroup::get_ID, ListViewItem::get_TileViewColumns, ExLVwLibU::TILEVIEWSUBITEM,
	///       ExplorerListView::get_hImageList, ExplorerListView::get_TileViewItemLines,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, MAX_ITEMTEXTLENGTH, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, ListViewItem::get_Text, ListViewItem::get_IconIndex,
	///       ListViewItem::get_Indent, ListViewItem::get_ItemData, ListViewItem::get_Group,
	///       ListViewGroup::get_ID, ListViewItem::get_TileViewColumns, ExLVwLibA::TILEVIEWSUBITEM,
	///       ExplorerListView::get_hImageList, ExplorerListView::get_TileViewItemLines,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, MAX_ITEMTEXTLENGTH, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR itemText, LONG insertAt = -1, LONG iconIndex = -2, LONG itemIndentation = 0, LONG itemData = 0, LONG groupID = -2, VARIANT tileViewColumns = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IListViewItem** ppAddedItem = NULL);
	/// \brief <em>Retrieves whether the specified item is part of the item collection</em>
	///
	/// \param[in] itemIdentifier A value that identifies the item to be checked.
	/// \param[in] groupIndex If the control is in virtual mode and \c itemIdentifier specifies the item's
	///            zero-based index, this parameter specifies the zero-based index of the group containing
	///            the item.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the item is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, ListViewItem::get_GroupIndex, ExplorerListView::get_VirtualMode, Add, Remove,
	///       ExLVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, ListViewItem::get_GroupIndex, ExplorerListView::get_VirtualMode, Add, Remove,
	///       ExLVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG itemIdentifier, LONG groupIndex = 0, ItemIdentifierTypeConstants itemIdentifierType = iitIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the items in the collection</em>
	///
	/// Retrieves the number of \c ListViewItem objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified item in the collection from the listview</em>
	///
	/// \param[in] itemIdentifier A value that identifies the listview item to be removed.
	/// \param[in] itemIdentifierType A value specifying the meaning of \c itemIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, ExLVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, ExLVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType = iitIndex);
	/// \brief <em>Removes all items in the collection from the listview</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Filter validation
	///
	//@{
	/// \brief <em>Validates a filter for a \c VARIANT_BOOL type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c VARIANT_BOOL.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidIntegerFilter, IsValidListViewGroupFilter, IsValidListViewWorkAreaFilter,
	///     IsValidStringFilter, put_Filter
	BOOL IsValidBooleanFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c SAFEARRAY (containing \c LONG values or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c SAFEARRAY
	/// (containing \c LONG values or compatible).
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerFilter, IsValidListViewGroupFilter,
	///     IsValidListViewWorkAreaFilter, IsValidStringFilter, put_Filter
	BOOL IsValidIntegerArrayFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c LONG or compatible.
	///
	/// \param[in] filter The \c VARIANT to check.
	/// \param[in] minValue The minimum value that the corresponding property would accept.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerArrayFilter, IsValidListViewGroupFilter,
	///     IsValidListViewWorkAreaFilter, IsValidStringFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c IListViewGroup type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c IListViewGroup.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerArrayFilter, IsValidIntegerFilter,
	///     IsValidListViewWorkAreaFilter, IsValidStringFilter, put_Filter
	BOOL IsValidListViewGroupFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c IListViewWorkArea type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c IListViewWorkArea.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerArrayFilter, IsValidIntegerFilter,
	///     IsValidListViewGroupFilter, IsValidStringFilter, put_Filter
	BOOL IsValidListViewWorkAreaFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c BSTR type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c BSTR.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerArrayFilter, IsValidIntegerFilter,
	///     IsValidListViewGroupFilter, IsValidListViewWorkAreaFilter, put_Filter
	BOOL IsValidStringFilter(VARIANT& filter);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the control's first item that might be in the collection</em>
	///
	/// \param[in] numberOfItems The number of listview items in the control.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return The item being the collection's first candidate. -1 if no item was found.
	///
	/// \sa GetNextItemToProcess, Count, RemoveAll, Next
	int GetFirstItemToProcess(int numberOfItems, HWND hWndLvw);
	/// \brief <em>Retrieves the control's next item that might be in the collection</em>
	///
	/// \param[in] previousItem The item at which to start the search for a new collection candidate.
	/// \param[in] numberOfItems The number of items in the control.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return The next item being a candidate for the collection. -1 if no item was found.
	///
	/// \sa GetFirstItemToProcess, Count, RemoveAll, Next
	int GetNextItemToProcess(int previousItem, int numberOfItems, HWND hWndLvw);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified boolean value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsIntegerArrayInSafeArray, IsIntegerInSafeArray, IsListViewGroupInSafeArray,
	///     IsStringInSafeArray, get_ComparisonFunction
	BOOL IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified integer array</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] pArray The integer array to search for.
	/// \param[in] elements The number of elements in the array specified by \c pArray.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the \c SAFEARRAY contains the array; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerInSafeArray, IsListViewGroupInSafeArray,
	///     IsStringInSafeArray, get_ComparisonFunction
	BOOL IsIntegerArrayInSafeArray(LPSAFEARRAY pSafeArray, PINT pArray, int elements, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified integer value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerArrayInSafeArray, IsListViewGroupInSafeArray,
	///     IsStringInSafeArray, get_ComparisonFunction
	BOOL IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified group</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] groupID The unique ID of the group to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the group; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerArrayInSafeArray, IsIntegerInSafeArray,
	///     IsStringInSafeArray, get_ComparisonFunction
	BOOL IsListViewGroupInSafeArray(LPSAFEARRAY pSafeArray, int groupID, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified \c BSTR value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerArrayInSafeArray, IsIntegerInSafeArray,
	///     IsListViewGroupInSafeArray, get_ComparisonFunction
	BOOL IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether an item is part of the collection (applying the filters)</em>
	///
	/// \param[in] itemIndex The item to check.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return \c TRUE, if the item is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(LVITEMINDEX& itemIndex, HWND hWndLvw = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	#ifdef INCLUDESHELLBROWSERINTERFACE
		/// \brief <em>Adds an item to the listview</em>
		///
		/// Adds an item with the specified properties at the specified position in the control and returns the
		/// the inserted item's zero-based index.
		///
		/// \param[in] pItemText The new item's caption text. The maximum number of characters in this text
		///            is \c MAX_ITEMTEXTLENGTH.
		/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted
		///            as the last item.
		/// \param[in] iconIndex The zero-based index of the item's icon in the control's \c ilSmall,
		///            \c ilLarge, \c ilExtraLarge and \c ilHighResolution imagelists. If set to -1, the
		///            control will fire the \c ItemGetDisplayInfo event each time this property's value is
		///            required. A value of -2 means 'not specified' and is valid if there's no imagelist
		///            associated with the control.
		/// \param[in] overlayIndex The one-based index of the item's overlay icon in the control's \c ilSmall,
		///            \c ilLarge, \c ilExtraLarge and \c ilHighResolution imagelists. If set to 0, the new
		///            item won't have an overlay icon.
		/// \param[in] itemIndentation The new item's indentation in 'Details' view in image widths. If set
		///            to 1, the item's indentation will be 1 image width; if set to 2, it will be 2 image
		///            widths and so on. If set to -1, the control will fire the \c ItemGetDisplayInfo event
		///            each time this property's value is required.
		/// \param[in] itemData A \c LPARAM value that will be associated with the item.
		/// \param[in] isGhosted If set to \c TRUE, the item will be drawn ghosted; otherwise not.
		/// \param[in] groupID The unique ID of the group that the new item will belong to. If set to \c -2,
		///            the item won't belong to any group. With comctl32.dll version 6.10 or higher, if set to
		///            -1, the control will fire the \c ItemGetDisplayInfo event each time this property's
		///            value is required. However, this feature seems to be broken in comctl32.dll, so that the
		///            event is raised only once per item.
		/// \param[in] pTileViewColumns An array of \c TILEVIEWSUBITEM structs which define the sub-items that
		///            will be displayed below the new item's text in 'Tiles' and 'Extended Tiles' view. If set
		///            to an empty array, no details will be displayed. If set to \c Empty, the control will
		///            fire the \c ItemGetDisplayInfo event each time this property's value is required.
		/// \param[in] setShellItemFlag If \c TRUE, the control will flag the item as a shell item; otherwise
		///            not.
		///
		/// \return The inserted item's zero-based index.
		///
		/// \remarks The \c groupID and \c pTileViewColumns parameters will be ignored if comctl32.dll is
		///          used in a version older than 6.0.
		///
		/// \if UNICODE
		///   \sa Count, Remove, RemoveAll, ListViewItem::get_Text, ListViewItem::get_IconIndex,
		///       ListViewItem::get_Indent, ListViewItem::get_ItemData, ListViewItem::get_Group,
		///       ListViewGroup::get_ID, ListViewItem::get_TileViewColumns, ExLVwLibU::TILEVIEWSUBITEM,
		///       ExplorerListView::get_hImageList, ExplorerListView::get_TileViewItemLines,
		///       ExplorerListView::Raise_ItemGetDisplayInfo, MAX_ITEMTEXTLENGTH, ExLVwLibU::ImageListConstants
		/// \else
		///   \sa Count, Remove, RemoveAll, ListViewItem::get_Text, ListViewItem::get_IconIndex,
		///       ListViewItem::get_Indent, ListViewItem::get_ItemData, ListViewItem::get_Group,
		///       ListViewGroup::get_ID, ListViewItem::get_TileViewColumns, ExLVwLibA::TILEVIEWSUBITEM,
		///       ExplorerListView::get_hImageList, ExplorerListView::get_TileViewItemLines,
		///       ExplorerListView::Raise_ItemGetDisplayInfo, MAX_ITEMTEXTLENGTH, ExLVwLibA::ImageListConstants
		/// \endif
		int Add(LPTSTR pItemText, int insertAt, int iconIndex, int overlayIndex, int itemIndentation, LPARAM itemData, BOOL isGhosted, int groupID, __in VARIANT* pTileViewColumns, BOOL setShellItemFlag);
	#else
		/// \brief <em>Adds an item to the listview</em>
		///
		/// Adds an item with the specified properties at the specified position in the control and returns the
		/// the inserted item's zero-based index.
		///
		/// \param[in] pItemText The new item's caption text. The maximum number of characters in this text
		///            is \c MAX_ITEMTEXTLENGTH.
		/// \param[in] insertAt The new item's zero-based index. If set to -1, the item will be inserted
		///            as the last item.
		/// \param[in] iconIndex The zero-based index of the item's icon in the control's \c ilSmall,
		///            \c ilLarge, \c ilExtraLarge and \c ilHighResolution imagelists. If set to -1, the
		///            control will fire the \c ItemGetDisplayInfo event each time this property's value is
		///            required. A value of -2 means 'not specified' and is valid if there's no imagelist
		///            associated with the control.
		/// \param[in] overlayIndex The one-based index of the item's overlay icon in the control's \c ilSmall,
		///            \c ilLarge, \c ilExtraLarge and \c ilHighResolution imagelists. If set to 0, the new
		///            item won't have an overlay icon.
		/// \param[in] itemIndentation The new item's indentation in 'Details' view in image widths. If set
		///            to 1, the item's indentation will be 1 image width; if set to 2, it will be 2 image
		///            widths and so on. If set to -1, the control will fire the \c ItemGetDisplayInfo event
		///            each time this property's value is required.
		/// \param[in] itemData A \c LPARAM value that will be associated with the item.
		/// \param[in] isGhosted If set to \c TRUE, the item will be drawn ghosted; otherwise not.
		/// \param[in] groupID The unique ID of the group that the new item will belong to. If set to \c -2,
		///            the item won't belong to any group. With comctl32.dll version 6.10 or higher, if set to
		///            -1, the control will fire the \c ItemGetDisplayInfo event each time this property's
		///            value is required. However, this feature seems to be broken in comctl32.dll, so that the
		///            event is raised only once per item.
		/// \param[in] pTileViewColumns An array of column indexes which specify the columns that will be
		///            used to display additional details below the new item's text in 'Tiles' view. If set
		///            to an empty array, no details will be displayed. If set to \c Empty, the control will
		///            fire the \c ItemGetDisplayInfo event each time this property's value is required.
		///
		/// \return The inserted item's zero-based index.
		///
		/// \remarks The \c groupID and \c pTileViewColumns parameters will be ignored if comctl32.dll is
		///          used in a version older than 6.0.
		///
		/// \if UNICODE
		///   \sa Count, Remove, RemoveAll, ListViewItem::get_Text, ListViewItem::get_IconIndex,
		///       ListViewItem::get_Indent, ListViewItem::get_ItemData, ListViewItem::get_Group,
		///       ListViewGroup::get_ID, ListViewItem::get_TileViewColumns, ExplorerListView::get_hImageList,
		///       ExplorerListView::get_TileViewItemLines,  ExplorerListView::Raise_ItemGetDisplayInfo,
		///       MAX_ITEMTEXTLENGTH, ExLVwLibU::ImageListConstants
		/// \else
		///   \sa Count, Remove, RemoveAll, ListViewItem::get_Text, ListViewItem::get_IconIndex,
		///       ListViewItem::get_Indent, ListViewItem::get_ItemData, ListViewItem::get_Group,
		///       ListViewGroup::get_ID, ListViewItem::get_TileViewColumns, ExplorerListView::get_hImageList,
		///       ExplorerListView::get_TileViewItemLines,  ExplorerListView::Raise_ItemGetDisplayInfo,
		///       MAX_ITEMTEXTLENGTH, ExLVwLibA::ImageListConstants
		/// \endif
		int Add(LPTSTR pItemText, int insertAt, int iconIndex, int overlayIndex, int itemIndentation, LPARAM itemData, BOOL isGhosted, int groupID, __in VARIANT* pTileViewColumns);
	#endif
	/// \brief <em>Shortens a filter as much as possible</em>
	///
	/// Optimizes a filter by detecting redundancies, tautologies and so on.
	///
	/// \param[in] filteredProperty The filter to optimize. Any of the values defined by the
	///            \c FilteredPropertyConstants enumeration is valid.
	///
	/// \sa put_Filter, put_FilterType
	void OptimizeFilter(FilteredPropertyConstants filteredProperty);
	#ifdef USE_STL
		/// \brief <em>Removes the specified items</em>
		///
		/// \param[in] itemsToRemove A vector containing all items to remove.
		/// \param[in] hWndLvw The listview window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveItems(std::list<int>& itemsToRemove, HWND hWndLvw);
	#else
		/// \brief <em>Removes the specified items</em>
		///
		/// \param[in] itemsToRemove A vector containing all items to remove.
		/// \param[in] hWndLvw The listview window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndLvw);
	#endif

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS 14
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS];

		/// \brief <em>The \c ExplorerListView object that owns this collection</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>Holds the last enumerated item</em>
		LVITEMINDEX lastEnumeratedItem;
		/// \brief <em>If \c TRUE, we must filter the items</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties()
		{
			caseSensitiveFilters = FALSE;
			pOwnerExLvw = NULL;
			lastEnumeratedItem.iItem = -1;
			lastEnumeratedItem.iGroup = 0;

			for(int i = 0; i < NUMBEROFFILTERS; ++i) {
				VariantInit(&filter[i]);
				filterType[i] = ftDeactivated;
				comparisonFunction[i] = NULL;
			}
			usingFilters = FALSE;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains the items in this collection.
		///
		/// \sa pOwnerExLvw
		HWND GetExLvwHWnd(void);
	} properties;

	/// \brief <em>Holds the \c CLSID of the \c TILEVIEWSUBITEM type</em>
	///
	/// \sa ListViewItems
	CLSID clsidTILEVIEWSUBITEM;
};     // ListViewItems

OBJECT_ENTRY_AUTO(__uuidof(ListViewItems), ListViewItems)