//////////////////////////////////////////////////////////////////////
/// \class ListViewItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing listview item</em>
///
/// This class is a wrapper around a listview item that - unlike an item wrapped by
/// \c VirtualListViewItem - really exists in the control.
///
/// \todo Improve documentation of the \c Hot property.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewItem, VirtualListViewItem, ListViewItems, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewItem, VirtualListViewItem, ListViewItems, ExplorerListView
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
#include "_IListViewItemEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE ListViewItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewItem, &CLSID_ListViewItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewItem>,
    public Proxy_IListViewItemEvents<ListViewItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListViewItem, &IID_IListViewItem, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewItem, &IID_IListViewItem, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWITEM)

		BEGIN_COM_MAP(ListViewItem)
			COM_INTERFACE_ENTRY(IListViewItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewItemEvents))
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
	/// \name Implementation of IListViewItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Activating property</em>
	///
	/// Retrieves whether the item is currently being activated. If this property is set to \c VARIANT_TRUE,
	/// the item is being activated; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this item state.
	///
	/// \sa put_Activating, ExplorerListView::get_ItemActivationMode, ExplorerListView::get_CallBackMask,
	///     ExplorerListView::Raise_ItemActivate
	virtual HRESULT STDMETHODCALLTYPE get_Activating(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Activating property</em>
	///
	/// Sets whether the item is currently being activated. If this property is set to \c VARIANT_TRUE,
	/// the item is being activated; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this item state.
	///
	/// \sa get_Activating, ExplorerListView::put_ItemActivationMode, ExplorerListView::put_CallBackMask,
	///     ExplorerListView::Raise_ItemActivate
	virtual HRESULT STDMETHODCALLTYPE put_Activating(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Anchor property</em>
	///
	/// Retrieves whether the item is the control's anchor item, i. e. it's the item with which
	/// range-selection begins. If it is the anchor item, this property is set to \c VARIANT_TRUE;
	/// otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Caret, get_Selected, ExplorerListView::get_MultiSelect,
	///     ExplorerListView::get_AnchorItem
	virtual HRESULT STDMETHODCALLTYPE get_Anchor(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the item is the control's caret item, i. e. it has the focus. If it is the
	/// caret item, this property is set to \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Anchor, get_Selected, ExplorerListView::get_CaretItem, ExplorerListView::get_CallBackMask
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilited property</em>
	///
	/// Retrieves whether the item is drawn as the target of a drag'n'drop operation, i. e. whether its
	/// background is highlighted. If this property is set to \c VARIANT_TRUE, the item is highlighted;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_DropHilitedItem, ExplorerListView::get_CallBackMask, get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_DropHilited(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Ghosted property</em>
	///
	/// Retrieves whether the item's icon is drawn semi-transparent. If this property is set to
	/// \c VARIANT_TRUE, the item's icon is drawn semi-transparent; otherwise it's drawn normal.
	/// Usually you make items ghosted if they're hidden or selected for a cut-paste-operation.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Ghosted, get_IconIndex, ExplorerListView::get_CallBackMask
	virtual HRESULT STDMETHODCALLTYPE get_Ghosted(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Ghosted property</em>
	///
	/// Sets whether the item's icon is drawn semi-transparent. If this property is set to
	/// \c VARIANT_TRUE, the item's icon is drawn semi-transparent; otherwise it's drawn normal.
	/// Usually you make items ghosted if they're hidden or selected for a cut-paste-operation.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Ghosted, put_IconIndex, ExplorerListView::put_CallBackMask
	virtual HRESULT STDMETHODCALLTYPE put_Ghosted(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Glowing property</em>
	///
	/// Retrieves whether the item is tagged as "glowing". A "glowing" item doesn't look different,
	/// but you can use custom draw to accentuate items that have this state.
	/// If this property is set to \c VARIANT_TRUE, the item is "glowing"; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this item state.
	///
	/// \sa put_Glowing, get_Selected, ExplorerListView::get_CallBackMask, ExplorerListView::Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE get_Glowing(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Glowing property</em>
	///
	/// Sets whether the item is tagged as "glowing". A "glowing" item doesn't look different,
	/// but you can use custom draw to accentuate items that have this state.
	/// If this property is set to \c VARIANT_TRUE, the item is "glowing"; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this item state.
	///
	/// \sa get_Glowing, put_Selected, ExplorerListView::put_CallBackMask, ExplorerListView::Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE put_Glowing(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Group property</em>
	///
	/// Retrieves the item's group. If \c NULL, the item doesn't belong to any group.
	///
	/// \param[out] ppGroup Receives the group's \c IListViewGroup implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa putref_Group, ExplorerListView::get_Groups
	virtual HRESULT STDMETHODCALLTYPE get_Group(IListViewGroup** ppGroup);
	/// \brief <em>Sets the \c Group property</em>
	///
	/// Sets the item's group. If set to \c NULL, the item doesn't belong to any group.
	///
	/// \param[in] pNewGroup The new group's \c IListViewGroup implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Due to limitations in Windows' \c SysListView32 implementation, changing this property from
	///          a valid group to \c Nothing won't have any effect.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_Group, ExplorerListView::get_Groups
	virtual HRESULT STDMETHODCALLTYPE putref_Group(IListViewGroup* pNewGroup);
	/// \brief <em>Retrieves the current setting of the \c GroupIndex property</em>
	///
	/// Retrieves the zero-based index of the group in which the item is displayed. In virtual mode, the same
	/// item can be in multiple groups. All copies of the item have the same text, the same icon index, the
	/// same selection state etc., but properties like the item rectangle are different for each copy. The
	/// \c GroupIndex property holds the zero-based index of the group that contains the copy wrapped by the
	/// object.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Index, ExplorerListView::get_VirtualMode
	virtual HRESULT STDMETHODCALLTYPE get_GroupIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Hot property</em>
	///
	/// Retrieves whether the item is the control's hot item. Usually the hot item is underlined and/or
	/// highlighted. If it is the hot item, this property is set to \c VARIANT_TRUE; otherwise it's
	/// set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_HotItem, ExplorerListView::get_ItemActivationMode,
	///     ExplorerListView::get_UnderlinedItems
	virtual HRESULT STDMETHODCALLTYPE get_Hot(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's icon in the control's \c ilSmall, \c ilLarge,
	/// \c ilExtraLarge and \c ilHighResolution imagelists. If set to -2, no icon is displayed for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, get_OverlayIndex, get_StateImageIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, get_OverlayIndex, get_StateImageIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the item's icon in the control's \c ilSmall, \c ilLarge, \c ilExtraLarge
	/// and \c ilHighResolution imagelists. If set to -1, the control will fire the \c ItemGetDisplayInfo
	/// event each time this property's value is required. If set to -2, no icon is displayed for this item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, ExplorerListView::Raise_ItemGetDisplayInfo,
	///       put_OverlayIndex, put_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, ExplorerListView::Raise_ItemGetDisplayInfo,
	///       put_OverlayIndex, put_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks An item's ID will never change.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, get_GroupIndex, ExLVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_Index, get_GroupIndex, ExLVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Indent property</em>
	///
	/// Retrieves the item's indentation in 'Details' view in image widths. If set to 1, the item's
	/// indentation will be 1 image width; if set to 2, it will be 2 image widths and so on. If set to -1,
	/// the control will fire the \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c ilSmall imagelist must be valid, if you want to use item indentations.
	///
	/// \if UNICODE
	///   \sa put_Indent, ExplorerListView::get_View, ExplorerListView::get_hImageList,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_Indent, ExplorerListView::get_View, ExplorerListView::get_hImageList,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Indent(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Indent property</em>
	///
	/// Sets the item's indentation in 'Details' view in image widths. If set to 1, the item's
	/// indentation will be 1 image width; if set to 2, it will be 2 image widths and so on. If set to -1,
	/// the control will fire the \c ItemGetDisplayInfo event each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c ilSmall imagelist must be valid, if you want to use item indentations.
	///
	/// \if UNICODE
	///   \sa get_Indent, ExplorerListView::put_View, ExplorerListView::put_hImageList,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_Indent, ExplorerListView::put_View, ExplorerListView::put_hImageList,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Indent(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing items changes other items' indexes, the index is the best
	///          (and fastest) option to identify an index.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, get_GroupIndex, ExLVwLibU::ItemIdentifierTypeConstants
	/// \else
	///   \sa get_ID, get_GroupIndex, ExLVwLibA::ItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value associated with the item. Use this property to associate any data
	/// with the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ItemData, ExplorerListView::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Sets the \c ItemData property</em>
	///
	/// Sets the \c LONG value associated with the item. Use this property to associate any data
	/// with the item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ItemData, ExplorerListView::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE put_ItemData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the one-based index of the item's overlay icon in the control's \c ilSmall, \c ilLarge,
	/// \c ilExtraLarge and \c ilHighResolution imagelists. An index of 0 means that no overlay is drawn for
	/// this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OverlayIndex, ExplorerListView::get_hImageList, ExplorerListView::get_CallBackMask,
	///       get_IconIndex, get_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_OverlayIndex, ExplorerListView::get_hImageList, ExplorerListView::get_CallBackMask,
	///       get_IconIndex, get_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
	/// \brief <em>Sets the \c OverlayIndex property</em>
	///
	/// Sets the one-based index of the item's overlay icon in the control's \c ilSmall, \c ilLarge,
	/// \c ilExtraLarge and \c ilHighResolution imagelists. An index of 0 means that no overlay is drawn for
	/// this item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OverlayIndex, ExplorerListView::put_hImageList, ExplorerListView::put_CallBackMask,
	///       put_IconIndex, put_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_OverlayIndex, ExplorerListView::put_hImageList, ExplorerListView::put_CallBackMask,
	///       put_IconIndex, put_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OverlayIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the item is drawn as a selected item, i. e. whether its background is
	/// highlighted. If this property is set to \c VARIANT_TRUE, the item is highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Selected, ExplorerListView::get_MultiSelect, get_Anchor, get_Caret, get_DropHilited,
	///     ExplorerListView::get_CallBackMask
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Selected property</em>
	///
	/// Sets whether the item is drawn as a selected item, i. e. whether its background is highlighted.
	/// If this property is set to \c VARIANT_TRUE, the item is highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Selected, ExplorerListView::get_MultiSelect, ExplorerListView::putref_CaretItem,
	///     ExplorerListView::put_CallBackMask
	virtual HRESULT STDMETHODCALLTYPE put_Selected(VARIANT_BOOL newValue);
	#ifdef INCLUDESHELLBROWSERINTERFACE
		/// \brief <em>Retrieves the current setting of the \c ShellListViewItemObject property</em>
		///
		/// Retrieves the \c ShellListViewItem object of this listview item from the attached \c ShellListView
		/// control.
		///
		/// \param[out] ppItem Receives the object's \c IDispatch implementation.
		///
		/// \return An \c HRESULT error code.
		///
		/// \remarks This property is read-only.
		virtual HRESULT STDMETHODCALLTYPE get_ShellListViewItemObject(IDispatch** ppItem);
	#endif
	/// \brief <em>Retrieves the current setting of the \c StateImageIndex property</em>
	///
	/// Retrieves the one-based index of the item's state image in the control's \c ilState imagelist. The
	/// state image is drawn next to the item and usually a checkbox.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_StateImageIndex, ExplorerListView::get_hImageList, ExplorerListView::get_CallBackMask,
	///       get_IconIndex, get_OverlayIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_StateImageIndex, ExplorerListView::get_hImageList, ExplorerListView::get_CallBackMask,
	///       get_IconIndex, get_OverlayIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Sets the \c StateImageIndex property</em>
	///
	/// Sets the one-based index of the item's state image in the control's \c ilState imagelist. The
	/// state image is drawn next to the item and usually a checkbox.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_StateImageIndex, ExplorerListView::put_hImageList, ExplorerListView::put_CallBackMask,
	///       put_IconIndex, put_OverlayIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_StateImageIndex, ExplorerListView::put_hImageList, ExplorerListView::put_CallBackMask,
	///       put_IconIndex, put_OverlayIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StateImageIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c SubItems property</em>
	///
	/// Retrieves a collection object wrapping all sub-items of this item.
	///
	/// \param[out] ppSubItems Receives the collection's \c IListViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ListViewItems
	virtual HRESULT STDMETHODCALLTYPE get_SubItems(IListViewSubItems** ppSubItems);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, MAX_ITEMTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo event
	/// each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, MAX_ITEMTEXTLENGTH, ExplorerListView::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c TileViewColumns property</em>
	///
	/// Retrieves an array of \c TILEVIEWSUBITEM structs which define the sub-items that will be displayed
	/// below the item's text in 'Tiles' and 'Extended Tiles' view. If set to an empty array, no details will
	/// be displayed. If set to \c Empty, the control will fire the \c ItemGetDisplayInfo event each time
	/// this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \if UNICODE
	///   \sa put_TileViewColumns, ExplorerListView::get_View, ExplorerListView::get_Columns,
	///       ExplorerListView::get_TileViewItemLines, get_Text, ExplorerListView::Raise_ItemGetDisplayInfo,
	///       ExLVwLibU::TILEVIEWSUBITEM
	/// \else
	///   \sa put_TileViewColumns, ExplorerListView::get_View, ExplorerListView::get_Columns,
	///       ExplorerListView::get_TileViewItemLines, get_Text, ExplorerListView::Raise_ItemGetDisplayInfo,
	///       ExLVwLibA::TILEVIEWSUBITEM
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TileViewColumns(VARIANT* pValue);
	/// \brief <em>Sets the \c TileViewColumns property</em>
	///
	/// Sets an array of \c TILEVIEWSUBITEM structs which define the sub-items that will be displayed
	/// below the item's text in 'Tiles' and 'Extended Tiles' view. If set to an empty array, no details will
	/// be displayed. If set to \c Empty, the control will fire the \c ItemGetDisplayInfo event each time
	/// this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \if UNICODE
	///   \sa get_TileViewColumns, ExplorerListView::put_View, ExplorerListView::get_Columns,
	///       ExplorerListView::put_TileViewItemLines, put_Text, ExplorerListView::Raise_ItemGetDisplayInfo,
	///       ExLVwLibU::TILEVIEWSUBITEM
	/// \else
	///   \sa get_TileViewColumns, ExplorerListView::put_View, ExplorerListView::get_Columns,
	///       ExplorerListView::put_TileViewItemLines, put_Text, ExplorerListView::Raise_ItemGetDisplayInfo,
	///       ExLVwLibA::TILEVIEWSUBITEM
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_TileViewColumns(VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c WorkArea property</em>
	///
	/// Retrieves the item's working area. If the control doesn't have multiple working areas, this
	/// property's value will be \c NULL.
	///
	/// \param[out] ppWorkArea Receives the working area's \c IListViewWorkArea implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The working area, that the item belongs to, depends on the item's position.\n
	///          This property is read-only.
	///
	/// \sa SetPosition, GetPosition, ExplorerListView::get_WorkAreas
	virtual HRESULT STDMETHODCALLTYPE get_WorkArea(IListViewWorkArea** ppWorkArea);

	/// \brief <em>Retrieves an imagelist containing the item's drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of this item.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The imagelist containing the drag image.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa ExplorerListView::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Ensures the item is visible</em>
	///
	/// Ensures that the item is visible, scrolling the control, if necessary.
	///
	/// \param[in] mustBeEntirelyVisible Specifies whether the item must be entirely visible. If set
	///            to \c VARIANT_TRUE, the control ensures that the item is entirely visible; otherwise
	///            it's enough if the item is partially visible.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsVisible
	virtual HRESULT STDMETHODCALLTYPE EnsureVisible(VARIANT_BOOL mustBeEntirelyVisible = VARIANT_TRUE);
	/// \brief <em>Finds the next item with the specified characteristics</em>
	///
	/// Retrieves the control's next item with the specified characteristics. The item represented by the
	/// \c ListViewItem object is used as the starting point, but is excluded from the search.
	///
	/// \param[in] searchMode A value specifying the meaning of the \c searchFor parameter. Any of the
	///            values defined by the \c SearchModeConstants enumeration is valid.
	/// \param[in] searchFor The criterion that the item must fulfill to be returned by this method. This
	///            parameter's format depends on the \c searchMode parameter:
	///            - \c smItemData An integer value.
	///            - \c smText A string value.
	///            - \c smPartialText A string value.
	///            - \c smNearestPosition An array containing two integer values. The first one specifies the
	///              x-coordinate, the second one the y-coordinate (both in pixels and relative to the
	///              control's upper-left corner).
	/// \param[in] searchDirection A value specifying the direction to search. Any of the values
	///            defined by the \c SearchDirectionConstants enumeration is valid. This parameter is
	///            ignored if the \c searchFor parameter is not set to \c smNearestPosition.
	/// \param[in] wrapAtLastItem If set to \c VARIANT_TRUE, the search will be continued with the first
	///            item if the last item is reached. This parameter is ignored if \c searchMode is set to
	///            \c smNearestPosition.
	/// \param[out] ppFoundItem Receives the matching item's \c IListViewItem implementation. \c NULL
	///             if no matching item was found.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::FindItem, ExLVwLibU::SearchModeConstants,
	///       ExLVwLibU::SearchDirectionConstants, ExplorerListView::Raise_FindVirtualItem
	/// \else
	///   \sa ExplorerListView::FindItem, ExLVwLibA::SearchModeConstants,
	///       ExLVwLibA::SearchDirectionConstants, ExplorerListView::Raise_FindVirtualItem
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE FindNextItem(SearchModeConstants searchMode, VARIANT searchFor, SearchDirectionConstants searchDirection = sdNoneSpecific, VARIANT_BOOL wrapAtLastItem = VARIANT_TRUE, IListViewItem** ppFoundItem = NULL);
	/// \brief <em>Retrieves the item's position</em>
	///
	/// Retrieves the item's position (in pixels) within the control's client area.
	///
	/// \param[in,out] pX The x-coordinate (in pixels) of the upper-left corner of the item's bounding
	///                rectangle relative to the control's upper-left corner.
	/// \param[in,out] pY The y-coordinate (in pixels) of the upper-left corner of the item's bounding
	///                rectangle relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetPosition, GetRectangle
	virtual HRESULT STDMETHODCALLTYPE GetPosition(OLE_XPOS_PIXELS* pX, OLE_YPOS_PIXELS* pY);
	/// \brief <em>Retrieves the bounding rectangle of either the item or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// item or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c ItemRectangleTypeConstants enumeration is valid.
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the control's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the control's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the control's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::ItemRectangleTypeConstants
	/// \else
	///   \sa ExLVwLibA::ItemRectangleTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(ItemRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Retrieves whether the item is visible</em>
	///
	/// \param[out] pVisible If \c VARIANT_TRUE, the item is visible; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Microsoft's implementation of this feature in \c SysListView32 seems to be a bit buggy,
	///          so expect this method to return the wrong value in some cases.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa EnsureVisible, GetPosition, SetPosition
	virtual HRESULT STDMETHODCALLTYPE IsVisible(VARIANT_BOOL* pVisible);
	/// \brief <em>Sets the item's info tip text</em>
	///
	/// Sets the item's info tip text. If the mouse cursor is located over the item while this method is
	/// called, the control displays a tool tip with the specified text. Otherwise the method does nothing.\n
	/// This method is for scenarios in which retrieving an item's info tip text is time-consuming. The
	/// client application would handle the \c ItemGetInfoTipText event by starting info tip extraction and
	/// returning an empty info tip text. After extracting the text asynchronously, it would call
	/// \c SetInfoTipText to display it.
	///
	/// \param[in] text The info tip text to display.
	/// \param[out] pSuccess If \c VARIANT_TRUE, the method succeeded (this also means that the mouse cursor
	///             is located over the item); otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The info tip text is not persisted.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa Raise_ItemGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE SetInfoTipText(BSTR text, VARIANT_BOOL* pSuccess);
	/// \brief <em>Sets the item's position</em>
	///
	/// Sets the item's position (in pixels) within the control's client area. This affects 'Icons',
	/// 'Small Icons', 'Tiles' and 'Extended Tiles' view only.
	///
	/// \param[in] x The x-coordinate (in pixels) of the upper-left corner of the item's bounding
	///              rectangle relative to the control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the upper-left corner of the item's bounding
	///              rectangle relative to the control's upper-left corner.
	///
	/// \sa GetPosition, GetRectangle, ExplorerListView::get_View
	virtual HRESULT STDMETHODCALLTYPE SetPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Starts editing the item's text</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::EndLabelEdit, ExplorerListView::get_hWndEdit,
	///     ExplorerListView::Raise_StartingLabelEditing, ExplorerListView::Raise_RenamingItem
	virtual HRESULT STDMETHODCALLTYPE StartLabelEditing(void);
	/// \brief <em>Updates the item</em>
	///
	/// Updates the item. If the \c IExplorerListView::AutoArrangeItems property is not set to \c aaiNone,
	/// this method causes the control to be arranged.
	///
	/// \sa ExplorerListView::get_AutoArrangeItems, ExplorerListView::RedrawItems
	virtual HRESULT STDMETHODCALLTYPE Update(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given item</em>
	///
	/// Attaches this object to a given item, so that the item's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] itemIndex The item to attach to.
	///
	/// \sa Detach
	void Attach(LVITEMINDEX& itemIndex);
	/// \brief <em>Detaches this object from an item</em>
	///
	/// Detaches this object from the item it currently wraps, so that it doesn't wrap any item anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the item wrapped by this object.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(LPLVITEM pSource);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualListViewItem* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LPLVITEM pTarget, HWND hWndLvw = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualListViewItem* pTarget);
	/// \brief <em>Sets the owner of this item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this item</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The index of the item wrapped by this object</em>
		LVITEMINDEX itemIndex;

		Properties()
		{
			pOwnerExLvw = NULL;
			itemIndex.iItem = -2;
			itemIndex.iGroup = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains this item.
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

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] itemIndex The item for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LVITEMINDEX& itemIndex, LPLVITEM pTarget, HWND hWndLvw = NULL);
};     // ListViewItem

OBJECT_ENTRY_AUTO(__uuidof(ListViewItem), ListViewItem)