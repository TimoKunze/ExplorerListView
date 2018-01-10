//////////////////////////////////////////////////////////////////////
/// \class ListViewColumn
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing listview column</em>
///
/// This class is a wrapper around a listview column that - unlike a column wrapped by
/// \c VirtualListViewColumn - really exists in the control.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewColumn, VirtualListViewColumn, ListViewColumns, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewColumn, VirtualListViewColumn, ListViewColumns, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "CWindowEx.h"
#include "_IListViewColumnEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE ListViewColumn : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewColumn, &CLSID_ListViewColumn>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewColumn>,
    public Proxy_IListViewColumnEvents<ListViewColumn>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListViewColumn, &IID_IListViewColumn, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewColumn, &IID_IListViewColumn, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewColumns;
	friend class ClassFactory;

public:
	/// \brief <em>The resource ID of Windows' default down arrow bitmap</em>
	///
	/// \sa IDB_SHELL32_UPARROW, put_BitmapHandle
	#define IDB_SHELL32_DOWNARROW	134
	/// \brief <em>The resource ID of Windows' default up arrow bitmap</em>
	///
	/// \sa IDB_SHELL32_DOWNARROW, put_BitmapHandle
	#define IDB_SHELL32_UPARROW		133

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWCOLUMN)

		BEGIN_COM_MAP(ListViewColumn)
			COM_INTERFACE_ENTRY(IListViewColumn)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewColumn)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewColumnEvents))
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
	/// \name Implementation of IListViewColumn
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Alignment property</em>
	///
	/// Retrieves the alignment of the column's caption and content. Any of the values defined by
	/// the \c AlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Alignment, get_Caption, ExLVwLibU::AlignmentConstants
	/// \else
	///   \sa put_Alignment, get_Caption, ExLVwLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(AlignmentConstants* pValue);
	/// \brief <em>Sets the \c Alignment property</em>
	///
	/// Sets the alignment of the column's caption and content. Any of the values defined by
	/// the \c AlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Alignment, put_Caption, ExLVwLibU::AlignmentConstants
	/// \else
	///   \sa get_Alignment, put_Caption, ExLVwLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Alignment(AlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BitmapHandle property</em>
	///
	/// Retrieves the handle of the column's bitmap. If set to NULL, the column's caption doesn't
	/// contain a bitmap.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c SortArrow property must be set to \c saNone if you want to display a bitmap.
	///
	/// \sa put_BitmapHandle, get_IconIndex, get_SortArrow, get_ImageOnRight
	virtual HRESULT STDMETHODCALLTYPE get_BitmapHandle(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c BitmapHandle property</em>
	///
	/// Sets the handle of the column's bitmap. If set to NULL, the column's caption doesn't
	/// contain a bitmap. If set to -1 (-2), Windows' default down-arrow (up-arrow) is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The bitmap does NOT get destroyed automatically. This also applies for the bitmaps that the
	///          method will load if this property is set to -1 or -2.\n
	///          The \c SortArrow property must be set to \c saNone if you want to display a bitmap.
	///
	/// \sa get_BitmapHandle, put_IconIndex, put_SortArrow, put_ImageOnRight
	virtual HRESULT STDMETHODCALLTYPE put_BitmapHandle(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c Caption property</em>
	///
	/// Retrieves the column's caption. The maximum number of characters in this text is specified by
	/// \c MAX_COLUMNCAPTIONLENGTH. If set to \c NULL, the control will fire the \c HeaderItemGetDisplayInfo
	/// event each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Caption, MAX_COLUMNCAPTIONLENGTH, ExplorerListView::Raise_HeaderItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Caption(BSTR* pValue);
	/// \brief <em>Sets the \c Caption property</em>
	///
	/// Sets the column's caption. The maximum number of characters in this text is specified by
	/// \c MAX_COLUMNCAPTIONLENGTH. If set to \c NULL, the control will fire the \c HeaderItemGetDisplayInfo
	/// event each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Due to a bug in Windows' \c SysHeader32 implementation, you should set the \c ColumnData
	///            property to a unique value if you want to set the \c Caption property to \c NULL or the
	///            \c IconIndex property to -1.
	///
	/// \sa get_Caption, MAX_COLUMNCAPTIONLENGTH, ExplorerListView::Raise_HeaderItemGetDisplayInfo,
	///     put_ColumnData, put_IconIndex
	virtual HRESULT STDMETHODCALLTYPE put_Caption(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the column header is the control's caret column header, i. e. it has the focus. If
	/// it is the caret column header, this property is set to \c VARIANT_TRUE; otherwise it's set to
	/// \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Selected, ExplorerListView::get_CaretColumn
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ColumnData property</em>
	///
	/// Retrieves the \c LONG value associated with the column. Use this property to associate any data
	/// with the column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ColumnData, ExplorerListView::Raise_FreeColumnData
	virtual HRESULT STDMETHODCALLTYPE get_ColumnData(LONG* pValue);
	/// \brief <em>Sets the \c ColumnData property</em>
	///
	/// Sets the \c LONG value associated with the column. Use this property to associate any data
	/// with the column.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Due to a bug in Windows' \c SysHeader32 implementation, you should set the \c ColumnData
	///            property to a unique value if you want to set the \c Caption property to \c NULL or the
	///            \c IconIndex property to -1.
	///
	/// \sa get_ColumnData, ExplorerListView::Raise_FreeColumnData, put_Caption, put_IconIndex
	virtual HRESULT STDMETHODCALLTYPE put_ColumnData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultWidth property</em>
	///
	/// Retrieves the column's default width in pixels. This value is used when auto-sizing the column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_DefaultWidth, get_MinimumWidth, get_IdealWidth, get_Width, get_Resizable,
	///     ExplorerListView::get_ResizableColumns
	virtual HRESULT STDMETHODCALLTYPE get_DefaultWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c DefaultWidth property</em>
	///
	/// Sets the column's default width in pixels. This value is used when auto-sizing the column.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_DefaultWidth, put_MinimumWidth, put_Width, put_Resizable,
	///     ExplorerListView::get_ResizableColumns
	virtual HRESULT STDMETHODCALLTYPE put_DefaultWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves the column's filter, which is displayed in the filterbar below the header control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.
	///
	/// \sa put_Filter, ListViewColumns::ClearAllFilters, ExplorerListView::get_ShowFilterBar
	virtual HRESULT STDMETHODCALLTYPE get_Filter(VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets the column's filter, which is displayed in the filterbar below the header control. If set
	/// to \c Empty, the filter is cleared. If set to an integer value, the filter edit control displays
	/// an error if the user tries to enter a non-integer value. On Windows Vista, if set to a date value,
	/// the filter edit control is specialized to accept date values only.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.
	///
	/// \sa get_Filter, ListViewColumns::ClearAllFilters, ExplorerListView::put_ShowFilterBar
	virtual HRESULT STDMETHODCALLTYPE put_Filter(VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c Height property</em>
	///
	/// Retrieves the column header's height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Width, get_Left, get_Top, GetDropDownRectangle
	virtual HRESULT STDMETHODCALLTYPE get_Height(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the column's icon in the control's \c ilHeader imagelist. If set to
	/// -1, the control will fire the \c HeaderItemGetDisplayInfo event each time this property's value is
	/// required. If set to -2, the column's caption doesn't contain an icon.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c SortArrow property must be set to \c saNone if you want to display an icon.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, get_BitmapHandle, get_SortArrow,
	///       get_ImageOnRight, get_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, get_BitmapHandle, get_SortArrow,
	///       get_ImageOnRight, get_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the column's icon in the control's \c ilHeader imagelist. If set to -1,
	/// the control will fire the \c HeaderItemGetDisplayInfo event each time this property's value is
	/// required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c SortArrow property must be set to \c saNone if you want to display an icon.
	///
	/// \attention Due to a bug in Windows' \c SysHeader32 implementation, you should set the \c ColumnData
	///            property to a unique value if you want to set the \c Caption property to \c NULL or the
	///            \c IconIndex property to -1.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, put_BitmapHandle, put_SortArrow,
	///       put_ImageOnRight, ExplorerListView::Raise_HeaderItemGetDisplayInfo, put_ColumnData,
	///       put_Caption, put_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, put_BitmapHandle, put_SortArrow,
	///       put_ImageOnRight, ExplorerListView::Raise_HeaderItemGetDisplayInfo, put_ColumnData,
	///       put_Caption, put_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A column's ID will never change.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, get_Position, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa get_Index, get_Position, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c IdealWidth property</em>
	///
	/// Retrieves the column's ideal width in pixels. The ideal width is the width, to which the column
	/// would be resized if it was auto-sized.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_DefaultWidth, get_MinimumWidth, get_Width, get_Resizable,
	///     ExplorerListView::get_ResizableColumns
	virtual HRESULT STDMETHODCALLTYPE get_IdealWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c ImageOnRight property</em>
	///
	/// Retrieves whether the column header's image is displayed to the right of the text. If this property
	/// property is set to \c VARIANT_TRUE, the image is displayed to the right, otherwise to the left of
	/// the column's caption text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the column's caption contains a bitmap as well as an icon, the bitmap always is
	///          displayed to the right of the text.
	///
	/// \sa put_ImageOnRight, get_BitmapHandle, get_IconIndex, get_Caption
	virtual HRESULT STDMETHODCALLTYPE get_ImageOnRight(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ImageOnRight property</em>
	///
	/// Sets whether the column header's image is displayed to the right of the text. If this property
	/// property is set to \c VARIANT_TRUE, the image is displayed to the right, otherwise to the left of
	/// the column's caption text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the column's caption contains a bitmap as well as an icon, the bitmap always is
	///          displayed to the right of the text.
	///
	/// \sa get_ImageOnRight, put_BitmapHandle, put_IconIndex, put_Caption
	virtual HRESULT STDMETHODCALLTYPE put_ImageOnRight(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing columns changes other columns' indexes, the index is the best
	///          (and fastest) option to identify a column.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, get_Position, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa get_ID, get_Position, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Left property</em>
	///
	/// Retrieves the distance between the header control's left border and the column's left border
	/// in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Position, get_Top, get_Width, get_Height, GetDropDownRectangle
	virtual HRESULT STDMETHODCALLTYPE get_Left(OLE_XPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Locale property</em>
	///
	/// Retrieves the unique ID of the locale to use when sorting by this column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The locale is used when sorting using the \c sobNumericIntText, \c sobNumericFloatText,
	///          \c sobCurrencyText or \c sobDateTimeText sorting criterion.
	///
	/// \if UNICODE
	///   \sa put_Locale, get_TextParsingFlags, ExplorerListView::SortItems, ExLVwLibU::SortByConstants
	/// \else
	///   \sa put_Locale, get_TextParsingFlags, ExplorerListView::SortItems, ExLVwLibA::SortByConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Locale(LONG* pValue);
	/// \brief <em>Sets the \c Locale property</em>
	///
	/// Sets the unique ID of the locale to use when sorting by this column.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The locale is used when sorting using the \c sobNumericIntText, \c sobNumericFloatText,
	///          \c sobCurrencyText or \c sobDateTimeText sorting criterion.
	///
	/// \if UNICODE
	///   \sa get_Locale, put_TextParsingFlags, ExplorerListView::SortItems, ExLVwLibU::SortByConstants
	/// \else
	///   \sa get_Locale, put_TextParsingFlags, ExplorerListView::SortItems, ExLVwLibA::SortByConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Locale(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumWidth property</em>
	///
	/// Retrieves the column's minimum width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The native list view control doesn't seem to allow a minimum width smaller than 30
	///          pixels.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_MinimumWidth, get_DefaultWidth, get_IdealWidth, get_Width, get_Resizable,
	///     ExplorerListView::get_ResizableColumns,	ExplorerListView::get_UseMinColumnWidths
	virtual HRESULT STDMETHODCALLTYPE get_MinimumWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MinimumWidth property</em>
	///
	/// Sets the column's minimum width in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The native list view control doesn't seem to allow a minimum width smaller than 30
	///          pixels.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_MinimumWidth, put_DefaultWidth, put_Width, put_Resizable,
	///     ExplorerListView::put_ResizableColumns,	ExplorerListView::put_UseMinColumnWidths
	virtual HRESULT STDMETHODCALLTYPE put_MinimumWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c OwnerDrawn property</em>
	///
	/// Retrieves whether the client application draws this column header itself. If set to \c VARIANT_TRUE,
	/// the control will fire the \c HeaderOwnerDrawItem event each time this column header must be drawn.
	/// If set to \c VARIANT_FALSE, the control will draw the column header itself. In this case drawing can
	/// still be customized using the \c HeaderCustomDraw event.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_OwnerDrawn, ExplorerListView::get_OwnerDrawn, ExplorerListView::Raise_HeaderOwnerDrawItem,
	///     ExplorerListView::Raise_HeaderCustomDraw
	virtual HRESULT STDMETHODCALLTYPE get_OwnerDrawn(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c OwnerDrawn property</em>
	///
	/// Sets whether the client application draws this column header itself. If set to \c VARIANT_TRUE,
	/// the control will fire the \c HeaderOwnerDrawItem event each time this column header must be drawn.
	/// If set to \c VARIANT_FALSE, the control will draw the column header itself. In this case drawing can
	/// still be customized using the \c HeaderCustomDraw event.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_OwnerDrawn, ExplorerListView::put_OwnerDrawn, ExplorerListView::Raise_HeaderOwnerDrawItem,
	///     ExplorerListView::Raise_HeaderCustomDraw
	virtual HRESULT STDMETHODCALLTYPE put_OwnerDrawn(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Position property</em>
	///
	/// Retrieves the column's current position as a zero-based index. The left-most column has position
	/// 0, the next one to the right has position 1 and so on.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Position, get_ID, get_Index, ListViewColumns::get_Positions,
	///       ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa put_Position, get_ID, get_Index, ListViewColumns::get_Positions,
	///       ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Position(LONG* pValue);
	/// \brief <em>Sets the \c Position property</em>
	///
	/// Sets the column's current position as a zero-based index. The left-most column has position
	/// 0, the next one to the right has position 1 and so on.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Position, ListViewColumns::put_Positions, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa get_Position, ListViewColumns::put_Positions, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Position(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Resizable property</em>
	///
	/// Retrieves whether the column's width can be changed by the user. If set to \c VARIANT_TRUE, the
	/// column's width can be changed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Resizable, get_Width, ExplorerListView::get_ResizableColumns
	virtual HRESULT STDMETHODCALLTYPE get_Resizable(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Resizable property</em>
	///
	/// Sets whether the column's width can be changed by the user. If set to \c VARIANT_TRUE, the
	/// column's width can be changed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Resizable, put_Width, ExplorerListView::put_ResizableColumns
	virtual HRESULT STDMETHODCALLTYPE put_Resizable(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the column is the control's selected column. The selected column is highlighted.
	/// If this property is set to \c VARIANT_TRUE, the column is the selected one; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Caret, ExplorerListView::get_SelectedColumn, get_SortArrow
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	#ifdef INCLUDESHELLBROWSERINTERFACE
		/// \brief <em>Retrieves the current setting of the \c ShellListViewColumnObject property</em>
		///
		/// Retrieves the \c ShellListViewColumn object of this listview column from the attached
		/// \c ShellListView control.
		///
		/// \param[out] ppColumn Receives the object's \c IDispatch implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		virtual HRESULT STDMETHODCALLTYPE get_ShellListViewColumnObject(IDispatch** ppColumn);
	#endif
	/// \brief <em>Retrieves the current setting of the \c ShowDropDownButton property</em>
	///
	/// Retrieves whether the column header contains a drop-down area. If the user clicks this drop-down
	/// area, the client may display a popup window which lets the user configure the column.\n
	/// If set to \c VARIANT_TRUE, a drop-down button is displayed in the column header; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_ShowDropDownButton, ExplorerListView::Raise_ColumnDropDown, get_SortArrow,
	///     GetDropDownRectangle
	virtual HRESULT STDMETHODCALLTYPE get_ShowDropDownButton(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDropDownButton property</em>
	///
	/// Sets whether the column header contains a drop-down area. If the user clicks this drop-down
	/// area, the client may display a popup window which lets the user configure the column.\n
	/// If set to \c VARIANT_TRUE, a drop-down button is displayed in the column header; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_ShowDropDownButton, ExplorerListView::Raise_ColumnDropDown, put_SortArrow,
	///     GetDropDownRectangle
	virtual HRESULT STDMETHODCALLTYPE put_ShowDropDownButton(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SortArrow property</em>
	///
	/// Retrieves the kind of sort arrow that is displayed next to the column's caption text. Any of the
	/// values defined by the \c SortArrowConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SortArrow, get_BitmapHandle, get_IconIndex, get_Caption, get_ShowDropDownButton,
	///       ExplorerListView::get_ClickableColumnHeaders, ExplorerListView::get_SelectedColumn,
	///       ExLVwLibU::SortArrowConstants
	/// \else
	///   \sa put_SortArrow, get_BitmapHandle, get_IconIndex, get_Caption, get_ShowDropDownButton,
	///       ExplorerListView::get_ClickableColumnHeaders, ExplorerListView::get_SelectedColumn,
	///       ExLVwLibA::SortArrowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SortArrow(SortArrowConstants* pValue);
	/// \brief <em>Sets the \c SortArrow property</em>
	///
	/// Sets the kind of sort arrow that is displayed next to the column's caption text. Any of the
	/// values defined by the \c SortArrowConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SortArrow, put_BitmapHandle, put_IconIndex, put_Caption, put_ShowDropDownButton,
	///       ExplorerListView::put_ClickableColumnHeaders, ExplorerListView::putref_SelectedColumn,
	///       ExLVwLibU::SortArrowConstants
	/// \else
	///   \sa get_SortArrow, put_BitmapHandle, put_IconIndex, put_Caption, put_ShowDropDownButton,
	///       ExplorerListView::put_ClickableColumnHeaders, ExplorerListView::putref_SelectedColumn,
	///       ExLVwLibA::SortArrowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SortArrow(SortArrowConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c StateImageIndex property</em>
	///
	/// Retrieves the one-based index of the column header's state image in the control's \c ilHeaderState
	/// imagelist. The state image is drawn next to the column header and usually a checkbox.\n
	/// If set to 0, the column's caption doesn't contain a state image.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows support only two different state images\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa put_StateImageIndex, ExplorerListView::get_hImageList, get_IconIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_StateImageIndex, ExplorerListView::get_hImageList, get_IconIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Sets the \c StateImageIndex property</em>
	///
	/// Sets the one-based index of the column header's state image in the control's \c ilHeaderState
	/// imagelist. The state image is drawn next to the column header and usually a checkbox.\n
	/// If set to 0, the column's caption doesn't contain a state image.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows support only two different state images\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa get_StateImageIndex, ExplorerListView::put_hImageList, put_IconIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_StateImageIndex, ExplorerListView::put_hImageList, put_IconIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StateImageIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c TextParsingFlags property</em>
	///
	/// Retrieves the options to apply when parsing the text contained in this column into a numerical or
	/// date value. The parsing results may be used when sorting by this column.
	///
	/// \param[in] parsingFunction Specifies the parsing function for which to retrieve the options. Any of
	///            the values defined by the \c TextParsingFunctionConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The parsing options are used when sorting using the \c sobNumericIntText,
	///          \c sobNumericFloatText, \c sobCurrencyText or \c sobDateTimeText sorting criterion.\n
	///          They are also used when sorting using the \c sobText criterion if a specific locale
	///          identifier has been set for this column.
	///
	/// \if UNICODE
	///   \sa put_TextParsingFlags, get_Locale, ExplorerListView::SortItems, ExLVwLibU::SortByConstants,
	///       ExLVwLibU::TextParsingFunctionConstants
	/// \else
	///   \sa put_TextParsingFlags, get_Locale, ExplorerListView::SortItems, ExLVwLibA::SortByConstants,
	///       ExLVwLibA::TextParsingFunctionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG* pValue);
	/// \brief <em>Sets the \c TextParsingFlags property</em>
	///
	/// Sets the options to apply when parsing the text contained in this column into a numerical or
	/// date value. The parsing results may be used when sorting by this column.
	///
	/// \param[in] parsingFunction Specifies the parsing function for which to set the options. Any of
	///            the values defined by the \c TextParsingFunctionConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The parsing options are used when sorting using the \c sobNumericIntText,
	///          \c sobNumericFloatText, \c sobCurrencyText or \c sobDateTimeText sorting criterion.\n
	///          They are also used when sorting using the \c sobText criterion if a specific locale
	///          identifier has been set for this column.
	///
	/// \if UNICODE
	///   \sa get_TextParsingFlags, put_Locale, ExplorerListView::SortItems, ExLVwLibU::SortByConstants,
	///       ExLVwLibU::TextParsingFunctionConstants
	/// \else
	///   \sa get_TextParsingFlags, put_Locale, ExplorerListView::SortItems, ExLVwLibA::SortByConstants,
	///       ExLVwLibA::TextParsingFunctionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Top property</em>
	///
	/// Retrieves the distance between the header control's top border and the column's top border
	/// in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Left, get_Height, get_Width, GetDropDownRectangle
	virtual HRESULT STDMETHODCALLTYPE get_Top(OLE_YPOS_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Width property</em>
	///
	/// Retrieves the column's width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Width, ExLVwLibU::AutoSizeConstants, get_MinimumWidth, get_DefaultWidth, get_IdealWidth,
	///       get_Left, get_Height, get_Top, GetDropDownRectangle, get_Resizable,
	///       ExplorerListView::get_ResizableColumns, ExplorerListView::get_ShowHeaderChevron,
	///       ExplorerListView::Raise_BeginColumnResizing, ExplorerListView::Raise_ResizingColumn,
	///       ExplorerListView::Raise_EndColumnResizing
	/// \else
	///   \sa put_Width, ExLVwLibA::AutoSizeConstants, get_MinimumWidth, get_DefaultWidth, get_IdealWidth,
	///       get_Left, get_Height, get_Top, GetDropDownRectangle, get_Resizable,
	///       ExplorerListView::get_ResizableColumns, ExplorerListView::get_ShowHeaderChevron,
	///       ExplorerListView::Raise_BeginColumnResizing, ExplorerListView::Raise_ResizingColumn,
	///       ExplorerListView::Raise_EndColumnResizing
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Width(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Width property</em>
	///
	/// Sets the column's width in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property may also be set to any of the values defined by the \c AutoSizeConstants
	///          enumeration.
	///
	/// \if UNICODE
	///   \sa get_Width, ExLVwLibU::AutoSizeConstants, put_MinimumWidth, put_DefaultWidth, get_IdealWidth,
	///       ExplorerListView::GetStringWidth, ExplorerListView::put_ColumnWidth, put_Resizable,
	///       ExplorerListView::put_ResizableColumns, ExplorerListView::put_ShowHeaderChevron,
	///       ExplorerListView::Raise_BeginColumnResizing, ExplorerListView::Raise_ResizingColumn,
	///       ExplorerListView::Raise_EndColumnResizing
	/// \else
	///   \sa get_Width, ExLVwLibA::AutoSizeConstants, put_MinimumWidth, put_DefaultWidth, get_IdealWidth,
	///       ExplorerListView::GetStringWidth, ExplorerListView::put_ColumnWidth, put_Resizable,
	///       ExplorerListView::put_ResizableColumns, ExplorerListView::put_ShowHeaderChevron,
	///       ExplorerListView::Raise_BeginColumnResizing, ExplorerListView::Raise_ResizingColumn,
	///       ExplorerListView::Raise_EndColumnResizing
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Width(OLE_XSIZE_PIXELS newValue);

	/// \brief <em>Retrieves an imagelist containing the column header's drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of this column's header.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the header control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the header control's upper-left corner.
	/// \param[out] phImageList The imagelist containing the drag image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa ExplorerListView::CreateLegacyHeaderDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the column header's drop-down area</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's upper-left corner) of the
	/// column header's drop-down area.
	///
	/// \param[out] pLeft Receives the x-coordinate (in pixels) of the upper-left corner of the drop-down
	///             area's bounding rectangle relative to the header control's upper-left corner.
	/// \param[out] pTop Receives the y-coordinate (in pixels) of the upper-left corner of the drop-down
	///             area's bounding rectangle relative to the header control's upper-left corner.
	/// \param[out] pRight Receives the x-coordinate (in pixels) of the lower-right corner of the drop-down
	///             area's bounding rectangle relative to the header control's upper-left corner.
	/// \param[out] pBottom Receives the y-coordinate (in pixels) of the lower-right corner of the drop-down
	///             area's bounding rectangle relative to the header control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_ShowDropDownButton, get_Left, get_Top, get_Height, get_Width
	virtual HRESULT STDMETHODCALLTYPE GetDropDownRectangle(OLE_XPOS_PIXELS* pLeft = NULL, OLE_YPOS_PIXELS* pTop = NULL, OLE_XPOS_PIXELS* pRight = NULL, OLE_YPOS_PIXELS* pBottom = NULL);
	/// \brief <em>Starts editing the column's filter</em>
	///
	/// \param[in] applyCurrentFilter Specifies whether to apply the edit control's content as new
	///            filter if the user is editing a column's filter when this method is called. If
	///            \c VARIANT_TRUE, the filter is applied; otherwise it is discarded.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.
	///
	/// \sa put_Filter, ExplorerListView::put_ShowFilterBar
	virtual HRESULT STDMETHODCALLTYPE StartFilterEditing(VARIANT_BOOL applyCurrentFilter = VARIANT_TRUE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given column</em>
	///
	/// Attaches this object to a given column, so that the column's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] columnIndex The column to attach to.
	///
	/// \sa Detach
	void Attach(int columnIndex);
	/// \brief <em>Detaches this object from a column</em>
	///
	/// Detaches this object from the column it currently wraps, so that it doesn't wrap any column anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the column wrapped by this object.
	///
	/// \param[in] pSourceHDITEM The source \c HDITEM structure from which to copy the settings.
	/// \param[in] pSourceLVCOLUMN The source \c LVCOLUMN structure from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(LPHDITEM pSourceHDITEM, LPLVCOLUMN pSourceLVCOLUMN);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualListViewColumn* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTargetHDITEM The target \c HDITEM structure to which to copy the settings.
	/// \param[in] pTargetLVCOLUMN The target \c LVCOLUMN structure to which to copy the settings.
	/// \param[in] hWndHeader The header window the method will work on.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LPHDITEM pTargetHDITEM, LPLVCOLUMN pTargetLVCOLUMN, HWND hWndHeader = NULL, HWND hWndLvw = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualListViewColumn* pTarget);
	/// \brief <em>Sets the owner of this column</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this column</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The index of the column wrapped by this object</em>
		int columnIndex;

		Properties()
		{
			pOwnerExLvw = NULL;
			columnIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains this column.
		///
		/// \sa pOwnerExLvw, GetHeaderHWnd
		HWND GetExLvwHWnd(void);
		/// \brief <em>Retrieves the window handle of the owning listview's header control</em>
		///
		/// \return The window handle of the header control that contains this column.
		///
		/// \sa pOwnerExLvw, GetExLvwHWnd
		HWND GetHeaderHWnd(void);
	} properties;

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] columnIndex The column for which to save the settings.
	/// \param[in] pTargetHDITEM The target \c HDITEM structure to which to copy the settings.
	/// \param[in] pTargetLVCOLUMN The target \c LVCOLUMN structure to which to copy the settings.
	/// \param[in] hWndHeader The header window the method will work on.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(int columnIndex, LPHDITEM pTargetHDITEM, LPLVCOLUMN pTargetLVCOLUMN, HWND hWndHeader = NULL, HWND hWndLvw = NULL);
};     // ListViewColumn

OBJECT_ENTRY_AUTO(__uuidof(ListViewColumn), ListViewColumn)