//////////////////////////////////////////////////////////////////////
/// \class VirtualListViewGroup
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing listview group</em>
///
/// This class is a wrapper around a listview group that does not yet or not anymore exist in the control.
///
/// \remarks Requires comctl32.dll version 6.0 or higher.
///
/// \if UNICODE
///   \sa ExLVwLibU::IVirtualListViewGroup, ListViewGroup, ExplorerListView
/// \else
///   \sa ExLVwLibA::IVirtualListViewGroup, ListViewGroup, ExplorerListView
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "ExLVwU.h"
#else
	#include "ExLVwA.h"
#endif
#include "_IVirtualListViewGroupEvents_CP.h"
#include "helpers.h"
#include "ExplorerListView.h"


class ATL_NO_VTABLE VirtualListViewGroup : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualListViewGroup, &CLSID_VirtualListViewGroup>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualListViewGroup>,
    public Proxy_IVirtualListViewGroupEvents<VirtualListViewGroup>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualListViewGroup, &IID_IVirtualListViewGroup, &LIBID_ExLVwLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualListViewGroup, &IID_IVirtualListViewGroup, &LIBID_ExLVwLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ExplorerListView;
	friend class ListViewGroup;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALLISTVIEWGROUP)

		BEGIN_COM_MAP(VirtualListViewGroup)
			COM_INTERFACE_ENTRY(IVirtualListViewGroup)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualListViewGroup)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualListViewGroupEvents))
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
	/// \name Implementation of IVirtualListViewGroup
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
	///          \c alCenter.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Text, get_DescriptionTextTop, get_DescriptionTextBottom, ExLVwLibU::AlignmentConstants,
	///       ExLVwLibU::GroupComponentConstants
	/// \else
	///   \sa get_Text, get_DescriptionTextTop, get_DescriptionTextBottom, ExLVwLibA::AlignmentConstants,
	///       ExLVwLibA::GroupComponentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(GroupComponentConstants groupComponent = gcHeader, AlignmentConstants* pValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the group will be or was the control's caret group, i. e. it will have or had the
	/// focus. If it will be or was the caret group, this property is set to \c VARIANT_TRUE; otherwise it's
	/// set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_SubsetLinkFocused, ExplorerListView::get_CaretGroup, ExplorerListView::Raise_GroupGotFocus,
	///     ExplorerListView::Raise_GroupLostFocus
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Collapsed property</em>
	///
	/// Retrieves whether the group will be or was collapsed. If this property is set to \c VARIANT_TRUE,
	/// the items in this group will be or were hidden; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Collapsible, ExplorerListView::Raise_CollapsedGroup, ExplorerListView::Raise_ExpandedGroup
	virtual HRESULT STDMETHODCALLTYPE get_Collapsed(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Collapsible property</em>
	///
	/// Retrieves whether the group will be or was collapsible. If this property is set to \c VARIANT_TRUE,
	/// the user can or could collapse this group; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Collapsed, ExplorerListView::Raise_CollapsedGroup, ExplorerListView::Raise_ExpandedGroup
	virtual HRESULT STDMETHODCALLTYPE get_Collapsible(VARIANT_BOOL* pValue);
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
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa MAX_GROUPDESCRIPTIONTEXTLENGTH, get_DescriptionTextTop, get_Text, get_SubtitleText,
	///     get_TaskText, get_SubsetLinkText, get_Alignment
	virtual HRESULT STDMETHODCALLTYPE get_DescriptionTextBottom(BSTR* pValue);
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
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa MAX_GROUPDESCRIPTIONTEXTLENGTH, get_DescriptionTextBottom, get_Text, get_SubtitleText,
	///     get_TaskText, get_SubsetLinkText, get_Alignment
	virtual HRESULT STDMETHODCALLTYPE get_DescriptionTextTop(BSTR* pValue);
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
	///          Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExplorerListView::get_hImageList, get_DescriptionTextTop, get_DescriptionTextBottom,
	///       ExLVwLibU::ImageListConstants
	/// \else
	///   \sa ExplorerListView::get_hImageList, get_DescriptionTextTop, get_DescriptionTextBottom,
	///       ExLVwLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this group.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A group's ID won't change except it is changed explicitly.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::GroupIdentifierTypeConstants
	/// \else
	///   \sa ExLVwLibA::GroupIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Selected property</em>
	///
	/// Retrieves whether the group will be or was drawn as a selected group, i. e. whether its background
	/// will be or was highlighted. If this property is set to \c VARIANT_TRUE, the group will be or was
	/// highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Caret, ExplorerListView::Raise_GroupSelectionChanged
	virtual HRESULT STDMETHODCALLTYPE get_Selected(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ShowHeader property</em>
	///
	/// Retrieves whether the group's header will be or was displayed. If this property is set to
	/// \c VARIANT_TRUE, the group will be or was displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ShowHeader(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Subseted property</em>
	///
	/// Retrieves whether the listview control will or has displayed only a subset of the group's items if
	/// the listview control is or was not large enough to display all items without scrolling. If this
	/// property is set to \c VARIANT_TRUE, the control will or has hidden some of the group's items and
	/// instead will or has displayed a link for displaying them; otherwise the control will or has always
	/// displayed all of the group's items.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_SubsetLinkText, ExplorerListView::get_MinItemRowsVisibleInGroups
	virtual HRESULT STDMETHODCALLTYPE get_Subseted(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c SubsetLinkFocused property</em>
	///
	/// Retrieves whether the group's subset link will have or had the keyboard focus. If this property is
	/// set to \c VARIANT_TRUE, the group's subset link will have or had the keyboard focus; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_Subseted, get_SubsetLinkText, get_Caret, ExplorerListView::get_MinItemRowsVisibleInGroups
	virtual HRESULT STDMETHODCALLTYPE get_SubsetLinkFocused(VARIANT_BOOL* pValue);
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
	/// \sa MAX_GROUPSUBSETLINKTEXTLENGTH, get_Subseted, get_SubsetLinkFocused, get_Text, get_SubtitleText,
	///     get_TaskText, get_DescriptionTextTop, get_DescriptionTextBottom
	virtual HRESULT STDMETHODCALLTYPE get_SubsetLinkText(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c SubtitleText property</em>
	///
	/// Retrieves the text displayed as the group's subtitle. The number of characters in this text is
	/// specified by \c MAX_GROUPSUBTITLETEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa MAX_GROUPSUBTITLETEXTLENGTH, get_Text, get_TaskText, get_DescriptionTextTop,
	///     get_DescriptionTextBottom, get_SubsetLinkText
	virtual HRESULT STDMETHODCALLTYPE get_SubtitleText(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c TaskText property</em>
	///
	/// Retrieves the text displayed as the group's task link. The number of characters in this text is
	/// specified by \c MAX_GROUPTASKTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa MAX_GROUPTASKTEXTLENGTH, get_Text, get_SubtitleText, get_DescriptionTextTop,
	///     get_DescriptionTextBottom, get_SubsetLinkText, ExplorerListView::Raise_GroupTaskLinkClick
	virtual HRESULT STDMETHODCALLTYPE get_TaskText(BSTR* pValue);
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
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Alignment, ExLVwLibU::GroupComponentConstants, get_SubtitleText, get_TaskText,
	///       get_DescriptionTextTop, get_DescriptionTextBottom, get_SubsetLinkText, get_IconIndex,
	///       ExplorerListView::get_GroupHeaderForeColor, ExplorerListView::get_GroupFooterForeColor
	/// \else
	///   \sa get_Alignment, ExLVwLibA::GroupComponentConstants, get_SubtitleText, get_TaskText,
	///       get_DescriptionTextTop, get_DescriptionTextBottom, get_SubsetLinkText, get_IconIndex,
	///       ExplorerListView::get_GroupHeaderForeColor, ExplorerListView::get_GroupFooterForeColor
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Text(GroupComponentConstants groupComponent = gcHeader, BSTR* pValue = NULL);
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
	HRESULT LoadState(__in PLVGROUP pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in PLVGROUP pTarget);
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
		/// \brief <em>A structure holding this group's settings</em>
		LVGROUP settings;

		Properties()
		{
			pOwnerExLvw = NULL;
			ZeroMemory(&settings, sizeof(settings));
		}

		~Properties();

	} properties;
};     // VirtualListViewGroup

OBJECT_ENTRY_AUTO(__uuidof(VirtualListViewGroup), VirtualListViewGroup)