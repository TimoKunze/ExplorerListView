//////////////////////////////////////////////////////////////////////
/// \class VirtualListViewWorkAreas
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c VirtualListViewWorkArea objects</em>
///
/// This class provides easy access to collections of \c VirtualListViewWorkArea objects. A
/// \c VirtualListViewWorkAreas object is used to group the working areas that do not yet exist in the
/// control.
///
/// \if UNICODE
///   \sa ExLVwLibU::IVirtualListViewWorkAreas, VirtualListViewWorkArea, ListViewWorkAreas,
///       ExplorerListView
/// \else
///   \sa ExLVwLibA::IVirtualListViewWorkAreas, VirtualListViewWorkArea, ListViewWorkAreas,
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
#include "_IVirtualListViewWorkAreasEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "VirtualListViewWorkArea.h"


class ATL_NO_VTABLE VirtualListViewWorkAreas : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualListViewWorkAreas, &CLSID_VirtualListViewWorkAreas>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualListViewWorkAreas>,
    public Proxy_IVirtualListViewWorkAreasEvents<VirtualListViewWorkAreas>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualListViewWorkAreas, &IID_IVirtualListViewWorkAreas, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualListViewWorkAreas, &IID_IVirtualListViewWorkAreas, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALLISTVIEWWORKAREAS)

		BEGIN_COM_MAP(VirtualListViewWorkAreas)
			COM_INTERFACE_ENTRY(IVirtualListViewWorkAreas)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualListViewWorkAreas)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualListViewWorkAreasEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the working areas</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c VirtualListViewWorkArea objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x working areas</em>
	///
	/// Retrieves the next \c numberOfMaxItems working areas from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of working areas the array identified by \c pItems
	///            can contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a working area's \c IVirtualListViewWorkArea implementation.
	/// \param[out] pNumberOfItemsReturned The number of working areas that actually were copied to the
	///             array identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, VirtualListViewWorkArea,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// working area in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x working areas</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip working areas.
	///
	/// \param[in] numberOfItemsToSkip The number of working areas to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IVirtualListViewWorkAreas
	///
	//@{
	/// \brief <em>Retrieves a \c VirtualListViewWorkArea object from the collection</em>
	///
	/// Retrieves a \c VirtualListViewWorkArea object from the collection that wraps the listview working
	/// area identified by \c workAreaIndex.
	///
	/// \param[in] workAreaIndex A value that identifies the listview working area to be retrieved.
	/// \param[out] ppWorkArea Receives the working area's \c IVirtualListViewWorkArea implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa VirtualListViewWorkArea
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG workAreaIndex, IVirtualListViewWorkArea** ppWorkArea);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c VirtualListViewWorkArea objects
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

	/// \brief <em>Counts the working areas in the collection</em>
	///
	/// Retrieves the number of \c VirtualListViewWorkArea objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given working area collection</em>
	///
	/// Attaches this object to a given working area collection, so that the collection can be managed
	/// using this object's methods.
	///
	/// \param[in] numberOfWorkAreas The number of working areas defined by the \c pWorkAreas parameter.
	/// \param[in] pWorkAreas The working areas to attach to.
	///
	/// \sa Detach
	void Attach(int numberOfWorkAreas, LPRECT pWorkAreas);
	/// \brief <em>Detaches this object from a working area collection</em>
	///
	/// Detaches this object from the working area collection it currently wraps, so that it doesn't wrap
	/// any working area collection anymore.
	///
	/// \sa Attach
	void Detach(void);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the number of working areas wrapped by this collection</em>
		int numberOfWorkAreas;
		/// \brief <em>Holds the working areas wrapped by this collection</em>
		LPRECT pWorkAreas;
		/// \brief <em>Holds the next enumerated working area</em>
		int nextEnumeratedWorkArea;

		Properties()
		{
			numberOfWorkAreas = 0;
			pWorkAreas = NULL;
			nextEnumeratedWorkArea = 0;
		}
	} properties;
};     // VirtualListViewWorkAreas

OBJECT_ENTRY_AUTO(__uuidof(VirtualListViewWorkAreas), VirtualListViewWorkAreas)