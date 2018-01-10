//////////////////////////////////////////////////////////////////////
/// \class ListViewWorkAreas
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListViewWorkArea objects</em>
///
/// This class provides easy access to collections of \c ListViewWorkArea objects. A \c ListViewWorkAreas
/// object is used to group the control's working areas.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewWorkAreas, ListViewWorkArea, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewWorkAreas, ListViewWorkArea, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IListViewWorkAreasEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "ListViewWorkArea.h"


class ATL_NO_VTABLE ListViewWorkAreas : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewWorkAreas, &CLSID_ListViewWorkAreas>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewWorkAreas>,
    public Proxy_IListViewWorkAreasEvents<ListViewWorkAreas>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListViewWorkAreas, &IID_IListViewWorkAreas, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewWorkAreas, &IID_IListViewWorkAreas, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWWORKAREAS)

		BEGIN_COM_MAP(ListViewWorkAreas)
			COM_INTERFACE_ENTRY(IListViewWorkAreas)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewWorkAreas)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewWorkAreasEvents))
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
	/// the \c ListViewWorkArea objects managed by this collection object.
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
	///                the pointer to a working area's \c IListViewWorkArea implementation.
	/// \param[out] pNumberOfItemsReturned The number of working areas that actually were copied to the
	///             array identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListViewWorkArea,
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
	/// \name Implementation of IListViewWorkAreas
	///
	//@{
	/// \brief <em>Retrieves a \c ListViewWorkArea object from the collection</em>
	///
	/// Retrieves a \c ListViewWorkArea object from the collection that wraps the listview working area
	/// identified by \c workAreaIndex.
	///
	/// \param[in] workAreaIndex A value that identifies the listview working area to be retrieved.
	/// \param[out] ppWorkArea Receives the working area's \c IListViewWorkArea implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ListViewWorkArea, Add, Remove
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG workAreaIndex, IListViewWorkArea** ppWorkArea);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListViewWorkArea objects
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

	/// \brief <em>Adds a working area to the listview</em>
	///
	/// Adds a working area with the specified properties at the specified position in the control and
	/// returns a \c ListViewWorkArea object wrapping the inserted working area.
	///
	/// \param[in] xLeft The x-coordinate (in pixels) of the left border of the new working area's
	///            bounding rectangle relative to the control's upper-left corner.
	/// \param[in] yTop The y-coordinate (in pixels) of the top border of the new working area's
	///            bounding rectangle relative to the control's upper-left corner.
	/// \param[in] xRight The x-coordinate (in pixels) of the right border of the new working area's
	///            bounding rectangle relative to the control's upper-left corner.
	/// \param[in] yBottom The y-coordinate (in pixels) of the bottom border of the new working area's
	///            bounding rectangle relative to the control's upper-left corner.
	/// \param[in] insertAt The new working area's zero-based index. If set to -1, the working area
	///            will be inserted as the last working area.
	/// \param[out] ppAddedWorkArea Receives the added working area's \c IListViewWorkArea implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The maximum number of working areas is 16.
	///
	/// \sa Count, Remove, RemoveAll, ListViewWorkArea::SetRectangle, LV_MAX_WORKAREAS
	virtual HRESULT STDMETHODCALLTYPE Add(OLE_XPOS_PIXELS xLeft, OLE_YPOS_PIXELS yTop, OLE_XPOS_PIXELS xRight, OLE_YPOS_PIXELS yBottom, LONG insertAt = -1, IListViewWorkArea** ppAddedWorkArea = NULL);
	/// \brief <em>Counts the working areas in the collection</em>
	///
	/// Retrieves the number of \c ListViewWorkArea objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified working area in the collection from the listview</em>
	///
	/// \param[in] workAreaIndex A value that identifies the listview working area to be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG workAreaIndex);
	/// \brief <em>Removes all working areas in the collection from the listview</em>
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
		/// \brief <em>Holds the next enumerated working area</em>
		int nextEnumeratedWorkArea;

		Properties()
		{
			pOwnerExLvw = NULL;
			nextEnumeratedWorkArea = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains the working areas in this collection.
		///
		/// \sa pOwnerExLvw
		HWND GetExLvwHWnd(void);
	} properties;
};     // ListViewWorkAreas

OBJECT_ENTRY_AUTO(__uuidof(ListViewWorkAreas), ListViewWorkAreas)