//////////////////////////////////////////////////////////////////////
/// \class VirtualListViewItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing listview item</em>
///
/// This class is a wrapper around a listview item that does not yet or not anymore exist in the control.
///
/// \if UNICODE
///   \sa ExLVwLibU::IVirtualListViewItem, ListViewItem, ExplorerListView
/// \else
///   \sa ExLVwLibA::IVirtualListViewItem, ListViewItem, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IVirtualListViewItemEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE VirtualListViewItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualListViewItem, &CLSID_VirtualListViewItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualListViewItem>,
    public Proxy_IVirtualListViewItemEvents< VirtualListViewItem>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualListViewItem, &IID_IVirtualListViewItem, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualListViewItem, &IID_IVirtualListViewItem, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewItem;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALLISTVIEWITEM)

		BEGIN_COM_MAP(VirtualListViewItem)
			COM_INTERFACE_ENTRY(IVirtualListViewItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualListViewItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualListViewItemEvents))
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
	/// \name Implementation of IVirtualListViewItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Activating property</em>
	///
	/// Retrieves whether the item will be or was being activated. If this property is set to
	/// \c VARIANT_TRUE, the item will be or was being activated; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this item state.\n
	///          This property is read-only.
	///
	/// \sa ExplorerListView::get_ItemActivationMode, ExplorerListView::Raise_ItemActivate
	virtual HRESULT STDMETHODCALLTYPE get_Activating(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the item will be or was the control's caret item, i. e. it will have or had the
	/// focus. If it will be or was the caret item, this property is set to \c VARIANT_TRUE; otherwise
	/// it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Selected, ExplorerListView::get_CaretItem
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DropHilited property</em>
	///
	/// Retrieves whether the item will be or was drawn as the target of a drag'n'drop operation, i. e.
	/// whether its background will be or was highlighted. If this property is set to \c VARIANT_TRUE,
	/// the item will be or was highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_DropHilitedItem, get_Selected
	virtual HRESULT STDMETHODCALLTYPE get_DropHilited(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Ghosted property</em>
	///
	/// Retrieves whether the item's icon will be or was drawn semi-transparent. If this property is
	/// set to \c VARIANT_TRUE, the item's icon will be or was drawn semi-transparent; otherwise it
	/// will be or was drawn normal. Usually you make items ghosted if they're hidden or selected for a
	/// cut-paste-operation.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_IconIndex
	virtual HRESULT STDMETHODCALLTYPE get_Ghosted(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Glowing property</em>
	///
	/// Retrieves whether the item will be or was tagged as "glowing". A "glowing" item doesn't look
	/// different, but you can use custom draw to accentuate items that have this state.
	/// If this property is set to \c VARIANT_TRUE, the item will be or was "glowing"; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this item state.\n
	///          This property is read-only.
	///
	/// \sa get_Selected, ExplorerListView::Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE get_Glowing(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Group property</em>
	///
	/// Retrieves the item's group. If \c Nothing, the item won't or didn't belong to any group.
	///
	/// \param[out] ppGroup Receives the group's \c IListViewGroup implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa ExplorerListView::get_Groups
	virtual HRESULT STDMETHODCALLTYPE get_Group(IListViewGroup** ppGroup);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the item's icon in the control's \c ilSmall, \c ilLarge,
	/// \c ilExtraLarge and \c ilHighResolution imagelists. If set to -1, the control will fire the
	/// \c ItemGetDisplayInfo event each time this property's value is required. If set to -2, no icon
	/// is displayed for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_hImageList, ExplorerListView::Raise_ItemGetDisplayInfo, get_OverlayIndex,
	///       get_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerListView::get_hImageList, ExplorerListView::Raise_ItemGetDisplayInfo, get_OverlayIndex,
	///       get_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
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
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_View, ExplorerListView::get_hImageList,
	///     ExplorerListView::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Indent(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index that will identify or has identified the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ListViewItems::Add
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::Raise_FreeItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the one-based index of the item's overlay icon in the control's \c ilSmall, \c ilLarge,
	/// \c ilExtraLarge and \c ilHighResolution imagelists. An index of 0 means that no overlay will be or
	/// was drawn for this item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_hImageList, get_IconIndex, get_StateImageIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerListView::get_hImageList, get_IconIndex, get_StateImageIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the item will be or was drawn as a selected item, i. e. whether its background
	/// will be or was highlighted. If this property is set to \c VARIANT_TRUE, the item will be or was
	/// highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_MultiSelect, get_DropHilited
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c StateImageIndex property</em>
	///
	/// Retrieves the one-based index of the item's state image in the control's \c ilState imagelist. The
	/// state image is drawn next to the item and usually a checkbox.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_hImageList, get_IconIndex, get_OverlayIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerListView::get_hImageList, get_IconIndex, get_OverlayIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo event
	/// each time this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa MAX_ITEMTEXTLENGTH, ExplorerListView::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c TileViewColumns property</em>
	///
	/// Retrieves an array of \c TILEVIEWSUBITEM structs which define the sub-items that will be or were
	/// displayed below the item's text in 'Tiles' and 'Extended Tiles' view. If set to an empty array, no
	/// details will be or were displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.\n
	///          Requires comctl32.dll version 6.0 or higher.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_View, ExplorerListView::get_Columns,
	///       ExplorerListView::get_TileViewItemLines, get_Text, ExLVwLibU::TILEVIEWSUBITEM
	/// \else
	///   \sa ExplorerListView::get_View, ExplorerListView::get_Columns,
	///       ExplorerListView::get_TileViewItemLines, get_Text, ExLVwLibA::TILEVIEWSUBITEM
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_TileViewColumns(VARIANT* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in LPLVITEM pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in LPLVITEM pTarget);
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
		/// \brief <em>A structure holding this item's settings</em>
		LVITEM settings;

		Properties()
		{
			pOwnerExLvw = NULL;
			ZeroMemory(&settings, sizeof(settings));
			settings.iItem = -1;
		}

		~Properties();
	} properties;
};     // VirtualListViewItem

OBJECT_ENTRY_AUTO(__uuidof(VirtualListViewItem), VirtualListViewItem)