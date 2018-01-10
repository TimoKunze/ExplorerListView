//////////////////////////////////////////////////////////////////////
/// \class ListViewGroup
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing listview group</em>
///
/// This class is a wrapper around a listview group that - unlike a group wrapped by
/// \c VirtualListViewGroup - really exists in the control.
///
/// \remarks Requires comctl32.dll version 6.0 or higher.
///
/// \if UNICODE
///   \sa ExLVwLibU::IListViewGroup, VirtualListViewGroup, ListViewGroups, ExplorerListView
/// \else
///   \sa ExLVwLibA::IListViewGroup, VirtualListViewGroup, ListViewGroups, ExplorerListView
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
#include "_IListViewGroupEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE ListViewGroup : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ListViewGroup, &CLSID_ListViewGroup>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ListViewGroup>,
    public Proxy_IListViewGroupEvents<ListViewGroup>, 
    #ifdef UNICODE
    	public IDispatchImpl<IListViewGroup, &IID_IListViewGroup, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IListViewGroup, &IID_IListViewGroup, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewGroups;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LISTVIEWGROUP)

		BEGIN_COM_MAP(ListViewGroup)
			COM_INTERFACE_ENTRY(IListViewGroup)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ListViewGroup)
			CONNECTION_POINT_ENTRY(__uuidof(_IListViewGroupEvents))
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
	/// \name Implementation of IListViewGroup
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Alignment property</em>
	///
	/// Retrieves the alignment of the group's header or footer text. Any of the values defined by
	/// the \c AlignmentConstants enumeration is valid.
	///
	/// \param[in] groupComponent The part of the group for which to retrieve the setting. Any of the
	///            values defined by the \c GroupComponentConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The group's description is displayed only if the \c Alignment(gcHeader) property is set to
	///          \c alCenter.
	///
	/// \if UNICODE
	///   \sa put_Alignment, get_Text, get_DescriptionTextTop, get_DescriptionTextBottom,
	///       ExLVwLibU::AlignmentConstants, ExLVwLibU::GroupComponentConstants
	/// \else
	///   \sa put_Alignment, get_Text, get_DescriptionTextTop, get_DescriptionTextBottom,
	///       ExLVwLibA::AlignmentConstants, ExLVwLibA::GroupComponentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(GroupComponentConstants groupComponent = gcHeader, AlignmentConstants* pValue = NULL);
	/// \brief <em>Sets the \c Alignment property</em>
	///
	/// Sets the alignment of the group's header or footer text. Any of the values defined by
	/// the \c AlignmentConstants enumeration is valid.
	///
	/// \param[in] groupComponent The part of the group for which to set the property. Any of the
	///            values defined by the \c GroupComponentConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The group's description is displayed only if the \c Alignment(gcHeader) property is set to
	///          \c alCenter.
	///
	/// \if UNICODE
	///   \sa get_Alignment, put_Text, put_DescriptionTextTop, put_DescriptionTextBottom,
	///       ExLVwLibU::AlignmentConstants, ExLVwLibU::GroupComponentConstants
	/// \else
	///   \sa get_Alignment, put_Text, put_DescriptionTextTop, put_DescriptionTextBottom,
	///       ExLVwLibA::AlignmentConstants, ExLVwLibA::GroupComponentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Alignment(GroupComponentConstants groupComponent = gcHeader, AlignmentConstants newValue = alLeft);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the group is the control's caret group, i. e. it has the focus. If it is the caret
	/// group, this property is set to \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Selected, get_SubsetLinkFocused, ExplorerListView::get_CaretGroup,
	///     ExplorerListView::Raise_GroupGotFocus, ExplorerListView::Raise_GroupLostFocus
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Collapsed property</em>
	///
	/// Retrieves whether the group is collapsed. If this property is set to \c VARIANT_TRUE, the items
	/// in this group are hidden; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Collapsed, get_Collapsible, ExplorerListView::Raise_CollapsedGroup,
	///     ExplorerListView::Raise_ExpandedGroup
	virtual HRESULT STDMETHODCALLTYPE get_Collapsed(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Collapsed property</em>
	///
	/// Sets whether the group is collapsed. If this property is set to \c VARIANT_TRUE, the items
	/// in this group are hidden; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Collapsed, put_Collapsible, ExplorerListView::Raise_CollapsedGroup,
	///     ExplorerListView::Raise_ExpandedGroup
	virtual HRESULT STDMETHODCALLTYPE put_Collapsed(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Collapsible property</em>
	///
	/// Retrieves whether the group can be collapsed. If this property is set to \c VARIANT_TRUE, the group
	/// can be collapsed by the user; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Collapsible, get_Collapsed, ExplorerListView::Raise_CollapsedGroup,
	///     ExplorerListView::Raise_ExpandedGroup
	virtual HRESULT STDMETHODCALLTYPE get_Collapsible(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Collapsible property</em>
	///
	/// Sets whether the group can be collapsed. If this property is set to \c VARIANT_TRUE, the group
	/// can be collapsed by the user; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Collapsible, put_Collapsed, ExplorerListView::Raise_CollapsedGroup,
	///     ExplorerListView::Raise_ExpandedGroup
	virtual HRESULT STDMETHODCALLTYPE put_Collapsible(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DescriptionTextBottom property</em>
	///
	/// Retrieves the bottom part of the text displayed as the group's description. The number of characters
	/// in this text is specified by \c MAX_GROUPDESCRIPTIONTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The description is displayed only if the \c Alignment(gcHeader) property is set to
	///          \c alCenter.\n
	///          The group's description isn't displayed if the group doesn't have an icon.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_DescriptionTextBottom, MAX_GROUPDESCRIPTIONTEXTLENGTH, get_DescriptionTextTop, get_Text,
	///     get_SubtitleText, get_TaskText, get_SubsetLinkText, get_Alignment
	virtual HRESULT STDMETHODCALLTYPE get_DescriptionTextBottom(BSTR* pValue);
	/// \brief <em>Sets the \c DescriptionTextBottom property</em>
	///
	/// Sets the bottom part of the text displayed as the group's description. The number of characters
	/// in this text is specified by \c MAX_GROUPDESCRIPTIONTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The description is displayed only if the \c Alignment(gcHeader) property is set to
	///          \c alCenter.\n
	///          The group's description isn't displayed if the group doesn't have an icon.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_DescriptionTextBottom, MAX_GROUPDESCRIPTIONTEXTLENGTH, put_DescriptionTextTop, put_Text,
	///     put_SubtitleText, put_TaskText, put_SubsetLinkText, put_Alignment
	virtual HRESULT STDMETHODCALLTYPE put_DescriptionTextBottom(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c DescriptionTextBottom property</em>
	///
	/// Retrieves the top part of the text displayed as the group's description. The number of characters
	/// in this text is specified by \c MAX_GROUPDESCRIPTIONTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The description is displayed only if the \c Alignment(gcHeader) property is set to
	///          \c alCenter.\n
	///          The group's description isn't displayed if the group doesn't have an icon.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_DescriptionTextTop, MAX_GROUPDESCRIPTIONTEXTLENGTH, get_DescriptionTextBottom, get_Text,
	///     get_SubtitleText, get_TaskText, get_SubsetLinkText, get_Alignment
	virtual HRESULT STDMETHODCALLTYPE get_DescriptionTextTop(BSTR* pValue);
	/// \brief <em>Sets the \c DescriptionTextBottom property</em>
	///
	/// Sets the top part of the text displayed as the group's description. The number of characters
	/// in this text is specified by \c MAX_GROUPDESCRIPTIONTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The description is displayed only if the \c Alignment(gcHeader) property is set to
	///          \c alCenter.\n
	///          The group's description isn't displayed if the group doesn't have an icon.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_DescriptionTextTop, MAX_GROUPDESCRIPTIONTEXTLENGTH, put_DescriptionTextBottom, put_Text,
	///     put_SubtitleText, put_TaskText, put_SubsetLinkText, put_Alignment
	virtual HRESULT STDMETHODCALLTYPE put_DescriptionTextTop(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c ListItems property</em>
	///
	/// Retrieves the first listview item in the group.
	///
	/// \param[out] ppItem Receives the item object's \c IListViewItem implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_ListItems, get_ItemCount, ListViewItem
	virtual HRESULT STDMETHODCALLTYPE get_FirstItem(IListViewItem** ppItem);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the group header's icon in the control's \c ilGroups imagelist. If
	/// set to -1, the group header doesn't contain an icon.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The group's description isn't displayed if the group doesn't have an icon.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, get_DescriptionTextTop,
	///       get_DescriptionTextBottom, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, ExplorerListView::get_hImageList, get_DescriptionTextTop,
	///       get_DescriptionTextBottom, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the group header's icon in the control's \c ilGroups imagelist. If set
	/// to -1, the group header doesn't contain an icon.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The group's description isn't displayed if the group doesn't have an icon.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, put_DescriptionTextTop,
	///       put_DescriptionTextBottom, ExLVwLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, ExplorerListView::put_hImageList, put_DescriptionTextTop,
	///       put_DescriptionTextBottom, ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this group.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A group's ID won't change except it is changed explicitly.
	///
	/// \if UNICODE
	///   \sa put_ID, get_Index, get_Position, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa put_ID, get_Index, get_Position, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Sets the \c ID property</em>
	///
	/// Sets an unique ID identifying this group.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A group's ID won't change except it is changed explicitly.
	///
	/// \if UNICODE
	///   \sa get_ID, put_Index, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa get_ID, put_Index, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ID(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this group.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Adding or removing groups changes other groups' indexes.\n
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, get_Position, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa get_ID, get_Position, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ItemCount property</em>
	///
	/// Retrieves the number of items contained in the group.
	///
	/// \param[in] visibleSubsetOnly If set to \c VARIANT_TRUE and the group is subseted, the number returned
	///            is the number of the currently visible items; otherwise it is the number of all items in
	///            the group.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only if virtual mode is disabled.
	///
	/// \sa put_ItemCount, get_FirstItem, ExplorerListView::get_VirtualMode,
	///     ExplorerListView::get_VirtualItemCount
	virtual HRESULT STDMETHODCALLTYPE get_ItemCount(VARIANT_BOOL visibleSubsetOnly = VARIANT_FALSE, LONG* pValue = NULL);
	/// \brief <em>Sets the \c ItemCount property</em>
	///
	/// Sets the number of items contained in the group.
	///
	/// \param[in] visibleSubsetOnly If set to \c VARIANT_TRUE and the group is subseted, the number returned
	///            is the number of the currently visible items; otherwise it is the number of all items in
	///            the group.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only if virtual mode is disabled.
	///
	/// \sa get_ItemCount, get_FirstItem, ExplorerListView::put_VirtualMode,
	///     ExplorerListView::put_VirtualItemCount
	virtual HRESULT STDMETHODCALLTYPE put_ItemCount(VARIANT_BOOL visibleSubsetOnly = VARIANT_FALSE, LONG newValue = 0);
	/// \brief <em>Retrieves the current setting of the \c ListItems property</em>
	///
	/// Retrieves a collection object wrapping the control's listview items in the group.
	///
	/// \param[out] ppItems Receives the collection object's \c IListViewItems implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_FirstItem, get_ItemCount, ListViewItems
	virtual HRESULT STDMETHODCALLTYPE get_ListItems(IListViewItems** ppItems);
	/// \brief <em>Retrieves the current setting of the \c Position property</em>
	///
	/// Retrieves the group's current position as a zero-based index. The top-most group has position
	/// 0, the next one to the bottom has position 1 and so on.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, get_Index, ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa get_ID, get_Index, ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Position(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the group is drawn as a selected group, i. e. whether its background is
	/// highlighted. If this property is set to \c VARIANT_TRUE, the group is highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Selected, get_Caret, ExplorerListView::get_CaretGroup,
	///     ExplorerListView::Raise_GroupSelectionChanged
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Selected property</em>
	///
	/// Sets whether the group is drawn as a selected group, i. e. whether its background is highlighted.
	/// If this property is set to \c VARIANT_TRUE, the group is highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Selected, put_Caret, ExplorerListView::putref_CaretGroup,
	///     ExplorerListView::Raise_GroupSelectionChanged
	virtual HRESULT STDMETHODCALLTYPE put_Selected(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowHeader property</em>
	///
	/// Retrieves whether the group's header is displayed. If this property is set to \c VARIANT_TRUE, the
	/// group's header is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_ShowHeader, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ShowHeader(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowHeader property</em>
	///
	/// Sets whether the group's header is displayed. If this property is set to \c VARIANT_TRUE, the
	/// group's header is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_ShowHeader, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ShowHeader(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Subseted property</em>
	///
	/// Retrieves whether the listview control displays only a subset of the group's items if the listview
	/// control is not large enough to display all items without scrolling. If this property is set to
	/// \c VARIANT_TRUE, the control may hide some of the group's items and instead display a link for
	/// displaying them; otherwise the control always displays all of the group's items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_Subseted, get_SubsetLinkText, ExplorerListView::get_MinItemRowsVisibleInGroups
	virtual HRESULT STDMETHODCALLTYPE get_Subseted(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Subseted property</em>
	///
	/// Sets whether the listview control displays only a subset of the group's items if the listview
	/// control is not large enough to display all items without scrolling. If this property is set to
	/// \c VARIANT_TRUE, the control may hide some of the group's items and instead display a link for
	/// displaying them; otherwise the control always displays all of the group's items.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Subseted, put_SubsetLinkText, ExplorerListView::put_MinItemRowsVisibleInGroups
	virtual HRESULT STDMETHODCALLTYPE put_Subseted(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SubsetLinkFocused property</em>
	///
	/// Retrieves whether the group's subset link has the keyboard focus. If this property is set to
	/// \c VARIANT_TRUE, the group's subset link has the keyboard focus; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_SubsetLinkFocused, get_Subseted, get_SubsetLinkText, get_Caret,
	///     ExplorerListView::get_MinItemRowsVisibleInGroups
	virtual HRESULT STDMETHODCALLTYPE get_SubsetLinkFocused(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SubsetLinkFocused property</em>
	///
	/// Sets whether the group's subset link has the keyboard focus. If this property is set to
	/// \c VARIANT_TRUE, the group's subset link has the keyboard focus; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_SubsetLinkFocused, put_Subseted, put_SubsetLinkText, put_Caret,
	///     ExplorerListView::put_MinItemRowsVisibleInGroups
	virtual HRESULT STDMETHODCALLTYPE put_SubsetLinkFocused(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SubsetLinkText property</em>
	///
	/// Retrieves the text displayed as a link below the group if not all items of the group are displayed.
	/// Clicking this link will display the remaining items. The maximum number of characters in this text is
	/// specified by \c MAX_GROUPSUBSETLINKTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_SubtitleText, MAX_GROUPSUBSETLINKTEXTLENGTH, get_Subseted, get_SubsetLinkFocused, get_Text,
	///     get_SubtitleText, get_TaskText, get_DescriptionTextTop, get_DescriptionTextBottom
	virtual HRESULT STDMETHODCALLTYPE get_SubsetLinkText(BSTR* pValue);
	/// \brief <em>Sets the \c SubsetLinkText property</em>
	///
	/// Sets the text displayed as a link below the group if not all items of the group are displayed.
	/// Clicking this link will display the remaining items. The maximum number of characters in this text is
	/// specified by \c MAX_GROUPSUBSETLINKTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_SubtitleText, MAX_GROUPSUBSETLINKTEXTLENGTH, put_Subseted, put_SubsetLinkFocused, put_Text,
	///     put_SubtitleText, put_TaskText, put_DescriptionTextTop, put_DescriptionTextBottom
	virtual HRESULT STDMETHODCALLTYPE put_SubsetLinkText(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c SubtitleText property</em>
	///
	/// Retrieves the text displayed as the group's subtitle. The number of characters in this text is
	/// specified by \c MAX_GROUPSUBTITLETEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_SubtitleText, MAX_GROUPSUBTITLETEXTLENGTH, get_Text, get_TaskText, get_DescriptionTextTop,
	///     get_DescriptionTextBottom, get_SubsetLinkText
	virtual HRESULT STDMETHODCALLTYPE get_SubtitleText(BSTR* pValue);
	/// \brief <em>Sets the \c SubtitleText property</em>
	///
	/// Sets the text displayed as the group's subtitle. The number of characters in this text is
	/// specified by \c MAX_GROUPSUBTITLETEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_SubtitleText, MAX_GROUPSUBTITLETEXTLENGTH, put_Text, put_TaskText, put_DescriptionTextTop,
	///     put_DescriptionTextBottom, put_SubsetLinkText
	virtual HRESULT STDMETHODCALLTYPE put_SubtitleText(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c TaskText property</em>
	///
	/// Retrieves the text displayed as the group's task link. The number of characters in this text is
	/// specified by \c MAX_GROUPTASKTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_TaskText, MAX_GROUPTASKTEXTLENGTH, get_Text, get_SubtitleText, get_DescriptionTextTop,
	///     get_DescriptionTextBottom, get_SubsetLinkText, ExplorerListView::Raise_GroupTaskLinkClick
	virtual HRESULT STDMETHODCALLTYPE get_TaskText(BSTR* pValue);
	/// \brief <em>Sets the \c TaskText property</em>
	///
	/// Sets the text displayed as the group's task link. The number of characters in this text is
	/// specified by \c MAX_GROUPTASKTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_TaskText, MAX_GROUPTASKTEXTLENGTH, put_Text, put_SubtitleText, put_DescriptionTextTop,
	///     put_DescriptionTextBottom, put_SubsetLinkText, ExplorerListView::Raise_GroupTaskLinkClick
	virtual HRESULT STDMETHODCALLTYPE put_TaskText(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the group's header or footer text. The number of characters in this text is not limited.
	///
	/// \param[in] groupComponent The part of the group for which to retrieve the setting. Any of the
	///            values defined by the \c GroupComponentConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Text, get_Alignment, ExLVwLibU::GroupComponentConstants, get_SubtitleText, get_TaskText,
	///       get_DescriptionTextTop, get_DescriptionTextBottom, get_SubsetLinkText, get_IconIndex,
	///       ExplorerListView::get_GroupHeaderForeColor, ExplorerListView::get_GroupFooterForeColor
	/// \else
	///   \sa put_Text, get_Alignment, ExLVwLibA::GroupComponentConstants, get_SubtitleText, get_TaskText,
	///       get_DescriptionTextTop, get_DescriptionTextBottom, get_SubsetLinkText, get_IconIndex,
	///       ExplorerListView::get_GroupHeaderForeColor, ExplorerListView::get_GroupFooterForeColor
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Text(GroupComponentConstants groupComponent = gcHeader, BSTR* pValue = NULL);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the group's header or footer text. The number of characters in this text is not limited.
	///
	/// \param[in] groupComponent The part of the group for which to set the property. Any of the
	///            values defined by the \c GroupComponentConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Text, put_Alignment, ExLVwLibU::GroupComponentConstants, put_SubtitleText, put_TaskText,
	///       put_DescriptionTextTop, put_DescriptionTextBottom, put_SubsetLinkText, put_IconIndex,
	///       ExplorerListView::put_GroupHeaderForeColor, ExplorerListView::put_GroupFooterForeColor
	/// \else
	///   \sa get_Text, put_Alignment, ExLVwLibA::GroupComponentConstants, put_SubtitleText, put_TaskText,
	///       put_DescriptionTextTop, put_DescriptionTextBottom, put_SubsetLinkText, put_IconIndex,
	///       ExplorerListView::put_GroupHeaderForeColor, ExplorerListView::put_GroupFooterForeColor
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Text(GroupComponentConstants groupComponent = gcHeader, BSTR newValue = L"");

	/// \brief <em>Retrieves the bounding rectangle of either the group or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// group or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c GroupRectangleTypeConstants enumeration is valid.
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
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::GroupRectangleTypeConstants
	/// \else
	///   \sa ExLVwLibA::GroupRectangleTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(GroupRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given group</em>
	///
	/// Attaches this object to a given group, so that the group's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] groupID The group to attach to.
	///
	/// \sa Detach
	void Attach(int groupID);
	/// \brief <em>Detaches this object from a group</em>
	///
	/// Detaches this object from the group it currently wraps, so that it doesn't wrap any group anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the group wrapped by this object.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(PLVGROUP pSource);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualListViewGroup* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(PLVGROUP pTarget, HWND hWndLvw = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualListViewGroup* pTarget);
	/// \brief <em>Sets the owner of this group</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerExLvw
	void SetOwner(__in_opt ExplorerListView* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ExplorerListView object that owns this group</em>
		///
		/// \sa SetOwner
		ExplorerListView* pOwnerExLvw;
		/// \brief <em>The unique ID of the group wrapped by this object</em>
		int groupID;

		Properties()
		{
			pOwnerExLvw = NULL;
			groupID = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning listview's window handle</em>
		///
		/// \return The window handle of the listview that contains this group.
		///
		/// \sa pOwnerExLvw
		HWND GetExLvwHWnd(void);
	} properties;

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] groupID The group for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndLvw The listview window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(int groupID, PLVGROUP pTarget, HWND hWndLvw = NULL);
};     // ListViewGroup

OBJECT_ENTRY_AUTO(__uuidof(ListViewGroup), ListViewGroup)