//////////////////////////////////////////////////////////////////////
/// \class ListViewWorkArea
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing listview working area</em>
///
/// This class is a wrapper around a listview working area that - unlike a working area wrapped by
/// \c VirtualListViewWorkArea - really exists in the control.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewWorkArea, VirtualListViewWorkArea, ListViewWorkAreas, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewWorkArea, VirtualListViewWorkArea, ListViewWorkAreas, ExplorerListView
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
#include "_IListViewWorkAreaEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE ListViewWorkArea : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewWorkArea, &CLSID_ListViewWorkArea>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewWorkArea>,
    public Proxy_IListViewWorkAreaEvents<ListViewWorkArea>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListViewWorkArea, &IID_IListViewWorkArea, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewWorkArea, &IID_IListViewWorkArea, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWWORKAREA)

		BEGIN_COM_MAP(ListViewWorkArea)
			COM_INTERFACE_ENTRY(IListViewWorkArea)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewWorkArea)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewWorkAreaEvents))
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
	/// \name Implementation of IListViewWorkArea
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this working area.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Adding or removing working areas changes other working areas' indexes.\n
	///          This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ListItems property</em>
	///
	/// Retrieves a collection object wrapping the control's listview items in the working area.
	///
	/// \param[out] ppItems Receives the collection object's \c IListViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ListViewItems
	virtual HRESULT STDMETHODCALLTYPE get_ListItems(IListViewItems** ppItems);

	/// \brief <em>Retrieves the working area's rectangle</em>
	///
	/// Retrieves the working area's bounding rectangle (in pixels) within the control's client area.
	///
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
	/// \sa SetRectangle
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Sets the working area's rectangle</em>
	///
	/// Sets the working area's bounding rectangle (in pixels) within the control's client area.
	///
	/// \param[in] xLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///            relative to the control's upper-left corner.
	/// \param[in] yTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///            relative to the control's upper-left corner.
	/// \param[in] xRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///            relative to the control's upper-left corner.
	/// \param[in] yBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///            relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetRectangle
	virtual HRESULT STDMETHODCALLTYPE SetRectangle(OLE_XPOS_PIXELS xLeft, OLE_YPOS_PIXELS yTop, OLE_XPOS_PIXELS xRight, OLE_YPOS_PIXELS yBottom);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given working area</em>
	///
	/// Attaches this object to a given working area, so that the working area's properties can be
	/// retrieved and set using this object's methods.
	///
	/// \param[in] workAreaIndex The working area to attach to.
	///
	/// \sa Detach
	void Attach(int workAreaIndex);
	/// \brief <em>Detaches this object from a working area</em>
	///
	/// Detaches this object from the working area it currently wraps, so that it doesn't wrap any
	/// working area anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this working area</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this working area</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The index of the working area wrapped by this object</em>
		int workAreaIndex;

		Properties()
		{
			pOwnerExLvw = NULL;
			workAreaIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains this working area.
		///
		/// \sa pOwnerExLvw
		HWND GetExLvwHWnd(void);
	} properties;
};     // ListViewWorkArea

OBJECT_ENTRY_AUTO(__uuidof(ListViewWorkArea), ListViewWorkArea)