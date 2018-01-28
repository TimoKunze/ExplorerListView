//////////////////////////////////////////////////////////////////////
/// \class ListViewFooterItem
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing listview footer item</em>
///
/// This class is a wrapper around a listview footer item that really exists in the control.
///
/// \remarks Requires comctl32.dll version 6.10 or higher.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewFooterItem, ListViewFooterItems, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewFooterItem, ListViewFooterItems, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "CWindowEx2.h"
#include "_IListViewFooterItemEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE ListViewFooterItem : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewFooterItem, &CLSID_ListViewFooterItem>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewFooterItem>,
    public Proxy_IListViewFooterItemEvents<ListViewFooterItem>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListViewFooterItem, &IID_IListViewFooterItem, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewFooterItem, &IID_IListViewFooterItem, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewFooterItems;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWFOOTERITEM)

		BEGIN_COM_MAP(ListViewFooterItem)
			COM_INTERFACE_ENTRY(IListViewFooterItem)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewFooterItem)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewFooterItemEvents))
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
	/// \name Implementation of IListViewFooterItem
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the footer item is the control's caret footer item, i. e. it has the focus. If it
	/// is the caret footer item, this property is set to \c VARIANT_TRUE; otherwise it's set to
	/// \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ExplorerListView::get_CaretFooterItem
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this footer item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Adding or removing footer items changes other footer items' indexes.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, ExLVwLibU::FooterItemIdentifierTypeConstants
	/// \else
	///   \sa get_ID, ExLVwLibA::FooterItemIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemData property</em>
	///
	/// Retrieves the \c LONG value associated with the footer item. Use this property to associate any data
	/// with the footer item.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks With current versions of comctl32.dll a footer item's associated data cannot be changed
	///          after the item has been inserted.\n
	///          With current versions of comctl32.dll the events \c FooterItemClick and
	///          \c FreeFooterItemData are not fired for footer items whose associated data is 0.\n
	///          This property is read-only.
	///
	/// \sa ListViewFooterItems::Add, ExplorerListView::Raise_FooterItemClick,
	///     ExplorerListView::Raise_FreeFooterItemData
	virtual HRESULT STDMETHODCALLTYPE get_ItemData(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the footer item's text. The maximum number of characters in this text is 4096.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_IconIndex, ListViewFooterItems::Add
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);

	/// \brief <em>Retrieves the bounding rectangle of the footer item</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of the footer
	/// item.
	///
	/// \param[out] xLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the control's upper-left corner.
	/// \param[out] yTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the control's upper-left corner.
	/// \param[out] xRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the control's upper-left corner.
	/// \param[out] yBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ExplorerListView::GetFooterRectangle
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given footer item</em>
	///
	/// Attaches this object to a given footer item, so that the footer item's properties can be retrieved
	/// and set using this object's methods.
	///
	/// \param[in] itemIndex The footer item to attach to.
	///
	/// \sa Detach
	void Attach(int itemIndex);
	/// \brief <em>Detaches this object from a footer item</em>
	///
	/// Detaches this object from the footer item it currently wraps, so that it doesn't wrap any footer item
	/// anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this footer item</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this footer item</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The zero-based index of the footer item wrapped by this object</em>
		int itemIndex;

		Properties()
		{
			pOwnerExLvw = NULL;
			itemIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains this footer item.
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
};     // ListViewFooterItem

OBJECT_ENTRY_AUTO(__uuidof(ListViewFooterItem), ListViewFooterItem)