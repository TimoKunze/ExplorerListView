//////////////////////////////////////////////////////////////////////
/// \class ListViewSubItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing listview sub-item</em>
///
/// This class is a wrapper around a listview sub-item.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewSubItem, ListViewSubItems, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewSubItem, ListViewSubItems, ExplorerListView
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
#include "_IListViewSubItemEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE ListViewSubItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewSubItem, &CLSID_ListViewSubItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewSubItem>,
    public Proxy_IListViewSubItemEvents<ListViewSubItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListViewSubItem, &IID_IListViewSubItem, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewSubItem, &IID_IListViewSubItem, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewSubItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWSUBITEM)

		BEGIN_COM_MAP(ListViewSubItem)
			COM_INTERFACE_ENTRY(IListViewSubItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewSubItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewSubItemEvents))
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
	/// \name Implementation of IListViewSubItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Activating property</em>
	///
	/// Retrieves whether the sub-item is currently being activated. If this property is set to
	/// \c VARIANT_TRUE, the sub-item is being activated; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this sub-item state.
	///
	/// \sa put_Activating, ExplorerListView::get_ItemActivationMode, ExplorerListView::Raise_ItemActivate
	virtual HRESULT STDMETHODCALLTYPE get_Activating(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Activating property</em>
	///
	/// Sets whether the sub-item is currently being activated. If this property is set to \c VARIANT_TRUE,
	/// the sub-item is being activated; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this sub-item state.
	///
	/// \sa get_Activating, ExplorerListView::get_ItemActivationMode, ExplorerListView::Raise_ItemActivate
	virtual HRESULT STDMETHODCALLTYPE put_Activating(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Ghosted property</em>
	///
	/// Retrieves whether the sub-item's icon is drawn semi-transparent. If this property is set to
	/// \c VARIANT_TRUE, the sub-item's icon is drawn semi-transparent; otherwise it's drawn normal.
	/// Usually you make items ghosted if they're hidden or selected for a cut-paste-operation.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Ghosted, get_IconIndex
	virtual HRESULT STDMETHODCALLTYPE get_Ghosted(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Ghosted property</em>
	///
	/// Sets whether the sub-item's icon is drawn semi-transparent. If this property is set to
	/// \c VARIANT_TRUE, the sub-item's icon is drawn semi-transparent; otherwise it's drawn normal.
	/// Usually you make items ghosted if they're hidden or selected for a cut-paste-operation.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Ghosted, put_IconIndex
	virtual HRESULT STDMETHODCALLTYPE put_Ghosted(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Glowing property</em>
	///
	/// Retrieves whether the sub-item is tagged as "glowing". A "glowing" sub-item doesn't look different,
	/// but you can use custom draw to accentuate sub-items that have this state.
	/// If this property is set to \c VARIANT_TRUE, the sub-item is "glowing"; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this sub-item state.
	///
	/// \sa put_Glowing, ExplorerListView::Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE get_Glowing(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Glowing property</em>
	///
	/// Sets whether the sub-item is tagged as "glowing". A "glowing" sub-item doesn't look different,
	/// but you can use custom draw to accentuate sub-items that have this state.
	/// If this property is set to \c VARIANT_TRUE, the sub-item is "glowing"; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Current versions of Windows do not use this sub-item state.
	///
	/// \sa get_Glowing, ExplorerListView::Raise_CustomDraw
	virtual HRESULT STDMETHODCALLTYPE put_Glowing(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the sub-item's icon in the control's \c ilSmall imagelist. If set
	/// to -1, the control will fire the \c ItemGetDisplayInfo event each time this property's value is
	/// required. If set to -2, no icon is displayed for this sub-item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, ExplorerListView::get_ShowSubItemImages,
	///       get_OverlayIndex, get_StateImageIndex, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, ExplorerListView::get_ShowSubItemImages,
	///       get_OverlayIndex, get_StateImageIndex, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the sub-item's icon in the control's \c ilSmall imagelist. If set to -1,
	/// the control will fire the \c ItemGetDisplayInfo event each time this property's value is required. If
	/// set to -2, no icon is displayed for this sub-item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, ExplorerListView::get_ShowSubItemImages,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, put_OverlayIndex, put_StateImageIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, ExplorerListView::get_ShowSubItemImages,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, put_OverlayIndex, put_StateImageIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves an one-based index identifying this sub-item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although there're no advantages from accessing an item through the \c ListViewSubItem
	///          class, it is possible. The sub-item with the index 0 is the parent item itself.\n
	///          Although adding or removing columns changes the sub-items' indexes, the index is the best
	///          (and fastest) option to identify a sub-item.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::ColumnIdentifierTypeConstants
	/// \else
	///   \sa ExLVwLibA::ColumnIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c OverlayIndex property</em>
	///
	/// Retrieves the zero-based index of the sub-item's overlay icon in the control's \c ilSmall imagelist.
	/// An index of 0 means that no overlay is drawn for this sub-item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OverlayIndex, ExplorerListView::get_hImageList, get_IconIndex, get_StateImageIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_OverlayIndex, ExplorerListView::get_hImageList, get_IconIndex, get_StateImageIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OverlayIndex(LONG* pValue);
	/// \brief <em>Sets the \c OverlayIndex property</em>
	///
	/// Sets the zero-based index of the sub-item's overlay icon in the control's \c ilSmall imagelist.
	/// An index of 0 means that no overlay is drawn for this sub-item.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OverlayIndex, ExplorerListView::put_hImageList, put_IconIndex, put_StateImageIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_OverlayIndex, ExplorerListView::put_hImageList, put_IconIndex, put_StateImageIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OverlayIndex(LONG newValue);
	/// \brief <em>Retrieves the sub-item's parent item</em>
	///
	/// \param[out] ppParentItem Receives the parent item's \c IListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ListViewItem
	virtual HRESULT STDMETHODCALLTYPE get_ParentItem(IListViewItem** ppParentItem);
	/// \brief <em>Retrieves the current setting of the \c StateImageIndex property</em>
	///
	/// Retrieves the one-based index of the sub-item's state image in the control's \c ilState imagelist.
	/// The state image is drawn next to the sub-item and usually a checkbox.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On current versions of Windows state images for sub-items are drawn, but misplaced and
	///          non-functional, so there's not much use for this property. Maybe this will change with
	///          future versions of Windows.
	///
	/// \if UNICODE
	///   \sa put_StateImageIndex, ExplorerListView::get_hImageList, get_IconIndex, get_OverlayIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_StateImageIndex, ExplorerListView::get_hImageList, get_IconIndex, get_OverlayIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_StateImageIndex(LONG* pValue);
	/// \brief <em>Sets the \c StateImageIndex property</em>
	///
	/// Sets the one-based index of the sub-item's state image in the control's \c ilState imagelist.
	/// The state image is drawn next to the sub-item and usually a checkbox.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks On current versions of Windows state images for sub-items are drawn, but misplaced and
	///          non-functional, so there's not much use for this property. Maybe this will change with
	///          future versions of Windows.
	///
	/// \if UNICODE
	///   \sa get_StateImageIndex, ExplorerListView::put_hImageList, put_IconIndex, put_OverlayIndex,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_StateImageIndex, ExplorerListView::put_hImageList, put_IconIndex, put_OverlayIndex,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_StateImageIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the sub-item's text. The maximum number of characters in this text is specified by
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
	/// Sets the sub-item's text. The maximum number of characters in this text is specified by
	/// \c MAX_ITEMTEXTLENGTH. If set to \c NULL, the control will fire the \c ItemGetDisplayInfo event
	/// each time this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, MAX_ITEMTEXTLENGTH, ExplorerListView::Raise_ItemGetDisplayInfo
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);

	/// \brief <em>Retrieves the bounding rectangle of either the sub-item or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// sub-item or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c ItemRectangleTypeConstants enumeration is valid.
	/// \param[out] PXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
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
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given sub-item</em>
	///
	/// Attaches this object to a given sub-item, so that the sub-item's properties can be retrieved
	/// and set using this object's methods.
	///
	/// \param[in] parentItemIndex The parent item of the sub-item to attach to.
	/// \param[in] subItemIndex The sub-item to attach to.
	///
	/// \sa Detach
	void Attach(LVITEMINDEX& parentItemIndex, int subItemIndex);
	/// \brief <em>Detaches this object from a sub-item</em>
	///
	/// Detaches this object from the sub-item it currently wraps, so that it doesn't wrap any sub-item
	/// anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this sub-item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this sub-item</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The index of the item the sub-item wrapped by this object belongs to</em>
		LVITEMINDEX parentItemIndex;
		/// \brief <em>The index of the sub-item wrapped by this object</em>
		int subItemIndex;

		Properties()
		{
			pOwnerExLvw = NULL;
			parentItemIndex.iItem = -1;
			parentItemIndex.iGroup = 0;
			subItemIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains this sub-item.
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
};     // ListViewSubItem

OBJECT_ENTRY_AUTO(__uuidof(ListViewSubItem), ListViewSubItem)