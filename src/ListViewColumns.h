//////////////////////////////////////////////////////////////////////
/// \class ListViewColumns
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListViewColumn objects</em>
///
/// This class provides easy access to collections of \c ListViewColumn objects. A \c ListViewColumns
/// object is used to group the control's columns.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewColumns, ListViewColumn, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewColumns, ListViewColumn, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IListViewColumnsEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "ListViewColumn.h"


class ATL_NO_VTABLE ListViewColumns : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewColumns, &CLSID_ListViewColumns>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewColumns>,
    public Proxy_IListViewColumnsEvents<ListViewColumns>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListViewColumns, &IID_IListViewColumns, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewColumns, &IID_IListViewColumns, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWCOLUMNS)

		BEGIN_COM_MAP(ListViewColumns)
			COM_INTERFACE_ENTRY(IListViewColumns)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewColumns)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewColumnsEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the columns</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ListViewColumn objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x columns</em>
	///
	/// Retrieves the next \c numberOfMaxItems columns from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of columns the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a column's \c IListViewColumn implementation.
	/// \param[out] pNumberOfItemsReturned The number of columns that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListViewColumn,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// column in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x columns</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip columns.
	///
	/// \param[in] numberOfItemsToSkip The number of columns to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IListViewColumns
	///
	//@{
	/// \brief <em>Retrieves a \c ListViewColumn object from the collection</em>
	///
	/// Retrieves a \c ListViewColumn object from the collection that wraps the listview column identified
	/// by \c columnIdentifier.
	///
	/// \param[in] columnIdentifier A value that identifies the listview column to be retrieved.
	/// \param[in] columnIdentifierType A value specifying the meaning of \c columnIdentifier. Any of the
	///            values defined by the \c ColumnIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppColumn Receives the column's \c IListViewColumn implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListViewColumn, Add, Remove, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa ListViewColumn, Add, Remove, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG columnIdentifier, ColumnIdentifierTypeConstants columnIdentifierType = citIndex, IListViewColumn** ppColumn = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListViewColumn objects
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
	/// \brief <em>Retrieves the current setting of the \c Positions property</em>
	///
	/// Retrieves the columns' positions. The sub-type of this property is an array of \c Longs. Each
	/// element of this array contains the index of the corresponding column, e. g. if the value of the
	/// 1st element is 2, the column, whose index is 2, is the first column in the control. The 2nd
	/// element represents the 2nd column and so on.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Positions, get_PositionsString, ListViewColumn::get_Position, ListViewColumn::get_Index
	virtual HRESULT STDMETHODCALLTYPE get_Positions(VARIANT* pValue);
	/// \brief <em>Sets the \c Positions property</em>
	///
	/// Sets the columns' positions. The sub-type of this property is an array of \c Longs. Each element
	/// of this array contains the index of the corresponding column, e. g. if the value of the 1st
	/// element is 2, the column, whose index is 2, is the first column in the control. The 2nd element
	/// represents the 2nd column and so on.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Positions, put_PositionsString, ListViewColumn::put_Position, ListViewColumn::get_Index
	virtual HRESULT STDMETHODCALLTYPE put_Positions(VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c PositionsString property</em>
	///
	/// Retrieves the columns' positions. This property works like the \c Positions property, but serializes
	/// the array into a string using the specified delimiter.
	///
	/// \param[in] stringDelimiter The string to use for delimiting the array elements when serializing
	///            it into a string.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_PositionsString, get_Positions, ListViewColumn::get_Position, ListViewColumn::get_Index
	virtual HRESULT STDMETHODCALLTYPE get_PositionsString(BSTR stringDelimiter, BSTR* pValue);
	/// \brief <em>Sets the \c PositionsString property</em>
	///
	/// Sets the columns' positions. This property works like the \c Positions property, but serializes
	/// the array into a string using the specified delimiter.
	///
	/// \param[in] stringDelimiter The string to use for delimiting the array elements when deserializing
	///            the string into an array.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_PositionsString, put_Positions, ListViewColumn::put_Position, ListViewColumn::get_Index
	virtual HRESULT STDMETHODCALLTYPE put_PositionsString(BSTR stringDelimiter, BSTR newValue);

	/// \brief <em>Adds a column to the listview</em>
	///
	/// Adds a column with the specified properties at the specified position in the control and returns a
	/// \c ListViewColumn object wrapping the inserted column.
	///
	/// \param[in] columnCaption The new column's caption. The maximum number of characters in this text
	///            is \c MAX_COLUMNCAPTIONLENGTH. If set to \c NULL, the control will fire the
	///            \c HeaderItemGetDisplayInfo event each time this property's value is required.
	/// \param[in] insertAt The new column's zero-based index. If set to -1, the column will be inserted
	///            as the last column.
	/// \param[in] columnWidth The new column's width in pixels.
	/// \param[in] minimumWidth The new column's minimum width in pixels.
	/// \param[in] alignment The alignment of the new column's caption and content. Any of the values
	///            defined by the \c AlignmentConstants enumeration is valid.
	/// \param[in] columnData A \c LONG value that will be associated with the column.
	/// \param[in] stateImageIndex The one-based index of the new column header's state image in the
	///            control's \c ilHeaderState imagelist. The state image is drawn next to the column header
	///            and usually a checkbox.\n
	///            If set to 0, the new column's caption doesn't contain a state image.
	/// \param[in] resizable If \c VARIANT_FALSE, the new column's width cannot be changed by the user.
	/// \param[in] showDropDownButton If \c VARIANT_TRUE, the new column's header contains a drop-down area.
	/// \param[in] ownerDrawn If \c VARIANT_TRUE, the new column's header will be owner-drawn, i. e. it will
	///            be drawn by the client application instead of the header control.
	/// \param[out] ppAddedColumn Receives the added column's \c IListViewColumn implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Due to a bug in Windows' \c SysHeader32 implementation, you should set the \c columnData
	///            parameter to a unique value if you want to set the \c columnCaption parameter to \c NULL
	///            or the inserted column's \c IconIndex property to -1.
	///
	/// \remarks The \c minimumWidth, \c stateImageIndex, \c resizable and \c showDropDownButton parameters
	///          are ignored if comctl32.dll is used in a version older than 6.10.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, ListViewColumn::get_Caption, ListViewColumn::get_Index,
	///       ListViewColumn::get_Width, ListViewColumn::get_MinimumWidth, ListViewColumn::get_Alignment,
	///       ListViewColumn::get_ColumnData, ListViewColumn::get_StateImageIndex,
	///       ExplorerListView::get_hImageList, ListViewColumn::get_Resizable,
	///       ListViewColumn::get_ShowDropDownButton, ListViewColumn::get_OwnerDrawn,
	///       ExplorerListView::Raise_HeaderItemGetDisplayInfo, ExLVwLibU::AlignmentConstants,
	///       MAX_COLUMNCAPTIONLENGTH, ListViewColumn::put_IconIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, ListViewColumn::get_Caption, ListViewColumn::get_Index,
	///       ListViewColumn::get_Width, ListViewColumn::get_MinimumWidth, ListViewColumn::get_Alignment,
	///       ListViewColumn::get_ColumnData, ListViewColumn::get_StateImageIndex,
	///       ExplorerListView::get_hImageList, ListViewColumn::get_Resizable,
	///       ListViewColumn::get_ShowDropDownButton, ListViewColumn::get_OwnerDrawn,
	///       ExplorerListView::Raise_HeaderItemGetDisplayInfo, ExLVwLibA::AlignmentConstants,
	///       MAX_COLUMNCAPTIONLENGTH, ListViewColumn::put_IconIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR columnCaption, LONG insertAt = -1, LONG columnWidth = 120, LONG minimumWidth = 0, AlignmentConstants alignment = alLeft, LONG columnData = 0, LONG stateImageIndex = 1, VARIANT_BOOL resizable = VARIANT_TRUE, VARIANT_BOOL showDropDownButton = VARIANT_FALSE, VARIANT_BOOL ownerDrawn = VARIANT_FALSE, IListViewColumn** ppAddedColumn = NULL);
	/// \brief <em>Clears all columns' filters</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.
	///
	/// \sa ListViewColumn::get_Filter, ExplorerListView::get_ShowFilterBar
	virtual HRESULT STDMETHODCALLTYPE ClearAllFilters(void);
	/// \brief <em>Counts the columns in the collection</em>
	///
	/// Retrieves the number of \c ListViewColumn objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified column in the collection from the listview</em>
	///
	/// \param[in] columnIdentifier A value that identifies the listview column to be removed.
	/// \param[in] columnIdentifierType A value specifying the meaning of \c columnIdentifier. Any of the
	///            values defined by the \c ItemIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Removing the column with the index 0 requires comctl32.dll version 5.0 or higher.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG columnIdentifier, ColumnIdentifierTypeConstants columnIdentifierType = citIndex);
	/// \brief <em>Removes all columns in the collection from the listview</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
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
		/// \brief <em>Holds the next enumerated column</em>
		int nextEnumeratedColumn;

		Properties()
		{
			pOwnerExLvw = NULL;
			nextEnumeratedColumn = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains the columns in this collection.
		///
		/// \sa pOwnerExLvw, GetHeaderHWnd
		HWND GetExLvwHWnd(void);
		/// \brief <em>Retrieves the window handle of the owning listview's header control</em>
		///
		/// \return The window handle of the header control that contains the columns in this collection.
		///
		/// \sa pOwnerExLvw, GetExLvwHWnd
		HWND GetHeaderHWnd(void);
	} properties;
};     // ListViewColumns

OBJECT_ENTRY_AUTO(__uuidof(ListViewColumns), ListViewColumns)