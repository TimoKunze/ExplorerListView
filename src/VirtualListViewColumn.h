//////////////////////////////////////////////////////////////////////
/// \class VirtualListViewColumn
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing listview column</em>
///
/// This class is a wrapper around a listview column that does not yet or not anymore exist in the control.
///
/// \if UNICODE
///   \sa ExLVwLibU::IVirtualListViewColumn, ListViewColumn, ExplorerListView
/// \else
///   \sa ExLVwLibA::IVirtualListViewColumn, ListViewColumn, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IVirtualListViewColumnEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE VirtualListViewColumn : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualListViewColumn, &CLSID_VirtualListViewColumn>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualListViewColumn>,
    public Proxy_IVirtualListViewColumnEvents<VirtualListViewColumn>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualListViewColumn, &IID_IVirtualListViewColumn, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualListViewColumn, &IID_IVirtualListViewColumn, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewColumn;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALLISTVIEWCOLUMN)

		BEGIN_COM_MAP(VirtualListViewColumn)
			COM_INTERFACE_ENTRY(IVirtualListViewColumn)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualListViewColumn)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualListViewColumnEvents))
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
	/// \name Implementation of IVirtualListViewColumn
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
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Caption, ExLVwLibU::AlignmentConstants
	/// \else
	///   \sa get_Caption, ExLVwLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(AlignmentConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c BitmapHandle property</em>
	///
	/// Retrieves the handle of the column's bitmap. If set to 0, the column's caption doesn't
	/// contain a bitmap.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The bitmap does NOT get destroyed automatically.\n
	///          The \c SortArrow property must be set to \c saNone if you want to display a bitmap.\n
	///          This property is read-only.
	///
	/// \sa get_IconIndex, get_SortArrow, get_ImageOnRight
	virtual HRESULT STDMETHODCALLTYPE get_BitmapHandle(OLE_HANDLE* pValue);
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
	/// \remarks This property is read-only.
	///
	/// \sa MAX_COLUMNCAPTIONLENGTH, ExplorerListView::Raise_HeaderItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Caption(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the column header will be or was the control's caret column header, i. e. it will
	/// have or had the focus. If it will be or was the caret column header, this property is set to
	/// \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa ExplorerListView::get_CaretColumn
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ColumnData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::Raise_FreeColumnData
	virtual HRESULT STDMETHODCALLTYPE get_ColumnData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c DefaultWidth property</em>
	///
	/// Retrieves the column's default width in pixels. This value is used when auto-sizing the column.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_MinimumWidth, get_Width, get_Resizable, ExplorerListView::get_ResizableColumns,
	///     ExplorerListView::get_ShowHeaderChevron
	virtual HRESULT STDMETHODCALLTYPE get_DefaultWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves the column's filter, which is displayed in the filterbar below the header control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.\n
	///          This property is read-only.
	///
	/// \sa ExplorerListView::get_ShowFilterBar
	virtual HRESULT STDMETHODCALLTYPE get_Filter(VARIANT* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the column's icon in the header control's \c ilHeader imagelist. If
	/// set to -1, the control will fire the \c HeaderItemGetDisplayInfo event each time this property's
	/// value is required. If set to -2, the column's caption doesn't contain an icon.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c SortArrow property must be set to \c saNone if you want to display an icon.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_hImageList, get_BitmapHandle, get_IconIndex, get_ImageOnRight,
	///       ExplorerListView::Raise_HeaderItemGetDisplayInfo, get_StateImageIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerListView::get_hImageList, get_BitmapHandle, get_IconIndex, get_ImageOnRight,
	///       ExplorerListView::Raise_HeaderItemGetDisplayInfo, get_StateImageIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
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
	///          displayed to the right of the text.\n
	///          This property is read-only.
	///
	/// \sa get_BitmapHandle, get_IconIndex, get_Caption
	virtual HRESULT STDMETHODCALLTYPE get_ImageOnRight(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumWidth property</em>
	///
	/// Retrieves the column's minimum width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_DefaultWidth, get_Width, get_Resizable, ExplorerListView::get_ResizableColumns,
	///     ExplorerListView::get_UseMinColumnWidths
	virtual HRESULT STDMETHODCALLTYPE get_MinimumWidth(OLE_XSIZE_PIXELS* pValue);
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
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_OwnerDrawn, ExplorerListView::Raise_HeaderOwnerDrawItem,
	///     ExplorerListView::Raise_HeaderCustomDraw
	virtual HRESULT STDMETHODCALLTYPE get_OwnerDrawn(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Position property</em>
	///
	/// Retrieves the column's current position as a zero-based index. The left-most column has position
	/// 0, the next one to the right has position 1 and so on.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListViewColumns::get_Positions, ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa ListViewColumns::get_Positions, ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Position(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Resizable property</em>
	///
	/// Retrieves whether the column's width can be changed by the user. If set to \c VARIANT_TRUE, the
	/// column's width can be changed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Width, ExplorerListView::get_ResizableColumns
	virtual HRESULT STDMETHODCALLTYPE get_Resizable(VARIANT_BOOL* pValue);
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
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa ExplorerListView::Raise_ColumnDropDown, get_SortArrow
	virtual HRESULT STDMETHODCALLTYPE get_ShowDropDownButton(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c SortArrow property</em>
	///
	/// Retrieves the kind of sort arrow that is displayed next to the column's caption text. Any of the
	/// values defined by the \c SortArrowConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_BitmapHandle, get_Caption, get_ShowDropDownButton,
	///       ExplorerListView::get_ClickableColumnHeaders, ExplorerListView::get_SelectedColumn,
	///       ExLVwLibU::SortArrowConstants
	/// \else
	///   \sa get_BitmapHandle, get_Caption, get_ShowDropDownButton,
	///       ExplorerListView::get_ClickableColumnHeaders, ExplorerListView::get_SelectedColumn,
	///       ExLVwLibA::SortArrowConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SortArrow(SortArrowConstants* pValue);
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
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_hImageList, get_IconIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerListView::get_hImageList, get_IconIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Width property</em>
	///
	/// Retrieves the column's width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MinimumWidth, get_DefaultWidth, get_Resizable, ExplorerListView::get_ResizableColumns,
	///     ExplorerListView::get_ShowHeaderChevron
	virtual HRESULT STDMETHODCALLTYPE get_Width(OLE_XSIZE_PIXELS* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSourceHDITEM The source \c HDITEM structure from which to copy the settings.
	/// \param[in] pSourceLVCOLUMN The source \c LVCOLUMN structure from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in LPHDITEM pSourceHDITEM, __in LPLVCOLUMN pSourceLVCOLUMN);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTargetHDITEM The target \c HDITEM structure to which to copy the settings.
	/// \param[in] pTargetLVCOLUMN The target \c LVCOLUMN structure to which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in LPHDITEM pTargetHDITEM, __in LPLVCOLUMN pTargetLVCOLUMN);
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
		/// \brief <em>A \c HDITEM structure holding this column header's settings</em>
		HDITEM headerSettings;
		/// \brief <em>A \c LVCOLUMN structure holding this column's settings</em>
		LVCOLUMN columnSettings;

		Properties()
		{
			pOwnerExLvw = NULL;
			ZeroMemory(&headerSettings, sizeof(headerSettings));
			ZeroMemory(&columnSettings, sizeof(columnSettings));
		}

		~Properties();

		/// \brief <em>Retrieves the window handle of the owning listview's header control</em>
		///
		/// \return The window handle of the header control that will contain or has contained this column.
		///
		/// \sa pOwnerExLvw
		HWND GetHeaderHWnd(void);
	} properties;
};     // VirtualListViewColumn

OBJECT_ENTRY_AUTO(__uuidof(VirtualListViewColumn), VirtualListViewColumn)