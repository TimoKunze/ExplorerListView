//////////////////////////////////////////////////////////////////////
/// \class VirtualListViewWorkArea
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing listview working area</em>
///
/// This class is a wrapper around a listview working area that does not yet exist in the control.
///
/// \if UNICODE
///   \sa ExLVwLibU::IVirtualListViewWorkArea, ListViewWorkArea, VirtualListViewWorkAreas,
///       ExplorerListView
/// \else
///   \sa ExLVwLibA::IVirtualListViewWorkArea, ListViewWorkArea, VirtualListViewWorkAreas,
///       ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IVirtualListViewWorkAreaEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE VirtualListViewWorkArea : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualListViewWorkArea, &CLSID_VirtualListViewWorkArea>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualListViewWorkArea>,
    public Proxy_IVirtualListViewWorkAreaEvents<VirtualListViewWorkArea>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualListViewWorkArea, &IID_IVirtualListViewWorkArea, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualListViewWorkArea, &IID_IVirtualListViewWorkArea, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALLISTVIEWWORKAREA)

		BEGIN_COM_MAP(VirtualListViewWorkArea)
			COM_INTERFACE_ENTRY(IVirtualListViewWorkArea)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualListViewWorkArea)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualListViewWorkAreaEvents))
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
	/// \name Implementation of IVirtualListViewWorkArea
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
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);

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
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given working area</em>
	///
	/// Attaches this object to a given working area, so that the working area's properties can be
	/// retrieved and set using this object's methods.
	///
	/// \param[in] pWorkAreas The control's working areas.
	/// \param[in] workAreaIndex The working area to attach to.
	///
	/// \sa Detach
	void Attach(LPRECT pWorkAreas, int workAreaIndex);
	/// \brief <em>Detaches this object from a working area</em>
	///
	/// Detaches this object from the working area it currently wraps, so that it doesn't wrap any
	/// working area anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the control's working areas</em>
		LPRECT pWorkAreas;
		/// \brief <em>The index of the working area wrapped by this object</em>
		int workAreaIndex;

		Properties()
		{
			workAreaIndex = -1;
			pWorkAreas = NULL;
		}
	} properties;
};     // VirtualListViewWorkArea

OBJECT_ENTRY_AUTO(__uuidof(VirtualListViewWorkArea), VirtualListViewWorkArea)