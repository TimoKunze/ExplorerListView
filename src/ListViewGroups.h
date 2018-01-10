//////////////////////////////////////////////////////////////////////
/// \class ListViewGroups
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ListViewGroup objects</em>
///
/// This class provides easy access to collections of \c ListViewGroup objects. A \c ListViewGroups
/// object is used to group the control's groups.
///
/// \remarks Requires comctl32.dll version 6.0 or higher.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewGroups, ListViewGroup, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewGroups, ListViewGroup, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IListViewGroupsEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"
#include "ListViewGroup.h"


class ATL_NO_VTABLE ListViewGroups : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewGroups, &CLSID_ListViewGroups>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewGroups>,
    public Proxy_IListViewGroupsEvents<ListViewGroups>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IListViewGroups, &IID_IListViewGroups, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewGroups, &IID_IListViewGroups, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWGROUPS)

		BEGIN_COM_MAP(ListViewGroups)
			COM_INTERFACE_ENTRY(IListViewGroups)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewGroups)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewGroupsEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the groups</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ListViewGroup objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x groups</em>
	///
	/// Retrieves the next \c numberOfMaxItems groups from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of groups the array identified by \c pItems can
	///            contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a group's \c IListViewGroup implementation.
	/// \param[out] pNumberOfItemsReturned The number of groups that actually were copied to the array
	///             identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ListViewGroup,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// group in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x groups</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip groups.
	///
	/// \param[in] numberOfItemsToSkip The number of groups to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IListViewGroups
	///
	//@{
	/// \brief <em>Retrieves a \c ListViewGroup object from the collection</em>
	///
	/// Retrieves a \c ListViewGroup object from the collection that wraps the listview group identified
	/// by \c groupIdentifier.
	///
	/// \param[in] groupIdentifier A value that identifies the listview group to be retrieved.
	/// \param[in] groupIdentifierType A value specifying the meaning of \c groupIdentifier. Any of the
	///            values defined by the \c GroupIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppGroup Receives the group's \c IListViewGroup implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ListViewGroup, Add, Remove, Contains, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa ListViewGroup, Add, Remove, Contains, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG groupIdentifier, GroupIdentifierTypeConstants groupIdentifierType = gitID, IListViewGroup** ppGroup = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ListViewGroup objects
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

	/// \brief <em>Adds a group to the listview</em>
	///
	/// Adds a group with the specified properties at the specified position in the control and returns a
	/// \c ListViewGroup object wrapping the inserted group.
	///
	/// \param[in] groupHeaderText The new group's header text. The number of characters in this text
	///            is not limited.
	/// \param[in] groupID The new group's ID. It must be unique and can't be set to -1.
	/// \param[in] insertAt The new group's zero-based index. If set to -1, the group will be inserted
	///            as the last group.
	/// \param[in] virtualItemCount If the listview control is in virtual mode, this parameter specifies
	///            the number of items in the new group.
	/// \param[in] headerAlignment The alignment of the new group's header text. Any of the values
	///            defined by the \c AlignmentConstants enumeration is valid.
	/// \param[in] iconIndex The zero-based index of the new group's header icon in the control's
	///            \c ilGroups imagelist. If set to -1, the group header doesn't contain an icon.
	/// \param[in] collapsible If \c VARIANT_TRUE, the new group can be collapsed by the user.
	/// \param[in] collapsed If \c VARIANT_TRUE, the new group is collapsed.
	/// \param[in] groupFooterText The new group's footer text. The number of characters in this text
	///            is not limited.
	/// \param[in] footerAlignment The alignment of the new group's footer text. Any of the values
	///            defined by the \c AlignmentConstants enumeration is valid.
	/// \param[in] subTitleText The new group's subtitle text. The maximum number of characters in this
	///            text is defined by \c MAX_GROUPSUBTITLETEXTLENGTH.
	/// \param[in] taskText The new group's task link text. The maximum number of characters in this
	///            text is defined by \c MAX_GROUPTASKTEXTLENGTH.
	/// \param[in] subsetLinkText The text displayed as a link below the new group if not all items of the
	///            group are displayed. The maximum number of characters in this text is defined by
	///            \c MAX_GROUPSUBSETLINKTEXTLENGTH.
	/// \param[in] subseted If \c VARIANT_TRUE, the control displays only a subset of the new group's items
	///            if the control is not large enough to display all items without scrolling.
	/// \param[in] showHeader If \c VARIANT_FALSE, the new group's header is not displayed.
	/// \param[out] ppAddedGroup Receives the added group's \c IListViewGroup implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c virtualItemCount, \c iconIndex, \c collapsible, \c collapsed, \c groupFooterText,
	///          \c footerAlignment, \c subTitleText, \c taskText, \c subsetLinkText, \c subseted and
	///          \c showHeader parameters are ignored if comctl32.dll is used in a version older than 6.10.\n
	///          The \c virtualItemCount parameter is ignored if the control is not in virtual mode.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, ListViewGroup::get_Text, ListViewGroup::get_ID,
	///       ListViewGroup::get_Position, ListViewGroup::get_ItemCount, ExplorerListView::get_VirtualMode,
	///       ListViewGroup::get_Alignment, ExLVwLibU::AlignmentConstants, ListViewGroup::get_IconIndex,
	///       ExplorerListView::get_hImageList, ExLVwLibU::ImageListConstants,
	///       ListViewGroup::get_Collapsible, ListViewGroup::get_Collapsed, ListViewGroup::get_SubtitleText,
	///       ListViewGroup::get_TaskText, ListViewGroup::get_SubsetLinkText, ListViewGroup::get_Subseted,
	///       ListViewGroup::get_ShowHeader
	/// \else
	///   \sa Count, Remove, RemoveAll, ListViewGroup::get_Text, ListViewGroup::get_ID,
	///       ListViewGroup::get_Position, ListViewGroup::get_ItemCount, ExplorerListView::get_VirtualMode,
	///       ListViewGroup::get_Alignment, ExLVwLibA::AlignmentConstants ListViewGroup::get_IconIndex,
	///       ExplorerListView::get_hImageList, ExLVwLibA::ImageListConstants,
	///       ListViewGroup::get_Collapsible, ListViewGroup::get_Collapsed, ListViewGroup::get_SubtitleText,
	///       ListViewGroup::get_TaskText, ListViewGroup::get_SubsetLinkText, ListViewGroup::get_Subseted,
	///       ListViewGroup::get_ShowHeader
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR groupHeaderText, LONG groupID, LONG insertAt = -1, LONG virtualItemCount = 1, AlignmentConstants headerAlignment = alLeft, LONG iconIndex = -1, VARIANT_BOOL collapsible = VARIANT_FALSE, VARIANT_BOOL collapsed = VARIANT_FALSE, BSTR groupFooterText = L"", AlignmentConstants footerAlignment = alLeft, BSTR subTitleText = L"", BSTR taskText = L"", BSTR subsetLinkText = L"", VARIANT_BOOL subseted = VARIANT_FALSE, VARIANT_BOOL showHeader = VARIANT_TRUE, IListViewGroup** ppAddedGroup = NULL);
	/// \brief <em>Retrieves whether the specified group is part of the group collection</em>
	///
	/// \param[in] groupIdentifier A value that identifies the group to be checked.
	/// \param[in] groupIdentifierType A value specifying the meaning of \c groupIdentifier. Any of the
	///            values defined by the \c GroupIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the group is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Remove, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa Add, Remove, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG groupIdentifier, GroupIdentifierTypeConstants groupIdentifierType = gitID, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the groups in the collection</em>
	///
	/// Retrieves the number of \c ListViewGroup objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified group in the collection from the listview</em>
	///
	/// \param[in] groupIdentifier A value that identifies the listview group to be removed.
	/// \param[in] groupIdentifierType A value specifying the meaning of \c groupIdentifier. Any of the
	///            values defined by the \c GroupIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG groupIdentifier, GroupIdentifierTypeConstants groupIdentifierType = gitID);
	/// \brief <em>Removes all groups in the collection from the listview</em>
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
		/// \brief <em>Holds the position index of the next enumerated group</em>
		int nextEnumeratedGroup;

		Properties()
		{
			pOwnerExLvw = NULL;
			nextEnumeratedGroup = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains the groups in this collection.
		///
		/// \sa pOwnerExLvw
		HWND GetExLvwHWnd(void);
	} properties;
};     // ListViewGroups

OBJECT_ENTRY_AUTO(__uuidof(ListViewGroups), ListViewGroups)