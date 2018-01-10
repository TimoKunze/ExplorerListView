//////////////////////////////////////////////////////////////////////
/// \file UndocComctl32.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Undocumented comctl32.dll stuff</em>
///
/// Declaration of some comctl32.dll internals that we're using.
///
/// \todo Check how far \c LVS_ALIGNBOTTOM and \c LVS_ALIGNRIGHT are implemented by Windows. It seems
///       to work fine on Windows XP SP2.
/// \todo Improve documentation ("See also" sections).
/// \todo Verify comctl32 requirements for \c LVN_GETEMPTYTEXT and \c LVM_RESETEMPTYTEXT.
/// \todo Check comctl32 requirements for \c LVM_GETHOTLIGHTCOLOR and \c LVM_SETHOTLIGHTCOLOR.
///
/// \sa ExplorerListView
//////////////////////////////////////////////////////////////////////

#pragma once


/// \brief <em>Retrieves a listview control's color for hot items' text</em>
///
/// This is an undocumented window message to retrieve the color that a listview will draw
/// hot items' text in.
///
/// \param[in] wParam Must be set to 0.
/// \param[in] lParam Must be set to 0.
///
/// \return The hot items' forecolor.
///
/// \sa LVM_SETHOTLIGHTCOLOR
#define LVM_GETHOTLIGHTCOLOR (LVM_FIRST + 79)
/// \brief <em>Sets a listview control's color for hot items' text</em>
///
/// This is an undocumented window message to set the color that a listview will draw hot items'
/// text in.
///
/// \param[in] wParam Must be set to 0.
/// \param[in] lParam The new color. May be \c CLR_DEFAULT.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \sa LVM_GETHOTLIGHTCOLOR
#define LVM_SETHOTLIGHTCOLOR (LVM_FIRST + 80)
/// \brief <em>Resets a listview control's emptiness text</em>
///
/// This is an undocumented window message to reset the text that a listview will display
/// when it's empty.
///
/// \param[in] wParam Must be set to 0.
/// \param[in] lParam Must be set to 0.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \remarks Requires comctl32.dll version 5.80 or higher.
///
/// \sa ExplorerListView::put_EmptyMarkupText
#define LVM_RESETEMPTYTEXT (LVM_FIRST + 84)
/// \brief <em>Queries a listview control for a COM interface</em>
///
/// This is an undocumented window message to query a COM interface from a listview control. Possible
/// interfaces are:
/// - \c IUnknown
/// - \c IOleWindow
/// - \c IListView
/// - \c IVisualProperties
/// - \c IPropertyControlSite
/// - \c IListViewFooter
///
/// \param[in] wParam Must be set to the address of an \c IID containing the requested interface's
///            interface ID.
/// \param[out] lParam Receives the address of the requested interface's implementation.
///
/// \return \c TRUE on success; otherwise \c FALSE.
///
/// \remarks Requires comctl32.dll version 6.10 or higher.
///
/// \sa IListViewFooter
#define LVM_QUERYINTERFACE (LVM_FIRST + 189)
/// \brief <em>Retrieves a listview control's emptiness text</em>
///
/// This is an undocumented notification message to retrieve the text that a listview will display
/// when it's empty.
///
/// \param[in] wParam The listview control's ID.
/// \param[in] lParam Address of a \c NMLVDISPINFOA struct, with \c mask set to \c LVIF_TEXT.
///
/// \return \c TRUE if the listview shall display a text.
///
/// \remarks Requires comctl32.dll version 5.80 or higher.
///
/// \sa ExplorerListView::OnGetEmptyTextNotification
#define LVN_GETEMPTYTEXTA (LVN_FIRST - 60)
/// \brief <em>Retrieves a listview control's emptiness text</em>
///
/// This is an undocumented notification message to retrieve the text that a listview will display
/// when it's empty.
///
/// \param[in] wParam The listview control's ID.
/// \param[in] lParam Address of a \c NMLVDISPINFOW struct, with \c mask set to \c LVIF_TEXT.
///
/// \return \c TRUE if the listview shall display a text.
///
/// \remarks Requires comctl32.dll version 5.80 or higher.
///
/// \sa ExplorerListView::OnGetEmptyTextNotification
#define LVN_GETEMPTYTEXTW (LVN_FIRST - 61)
#ifdef UNICODE
	#define LVN_GETEMPTYTEXT LVN_GETEMPTYTEXTW
#else
	#define LVN_GETEMPTYTEXT LVN_GETEMPTYTEXTA
#endif
/// \brief <em>Notifies a listview control's owner that a group's state has changed</em>
///
/// This is an undocumented notification message to notify a listview control's owner window that a group's
/// state has been changed, e.g. if the group has been collapsed.
///
/// \param[in] wParam The listview control's ID.
/// \param[in] lParam Address of a \c NMLVGROUP struct.
///
/// \return The return value seems to be ignored.
///
/// \remarks Requires comctl32.dll version 6.10 or higher.
///
/// \sa NMLVGROUP, ExplorerListView::OnGroupInfoNotification
#define LVN_GROUPINFO (LVN_FIRST - 88)

typedef struct tagNMLVGROUP
{
	NMHDR hdr;
	int iGroupId;
	UINT uNewState;
	UINT uOldState;
} NMLVGROUP, *LPNMLVGROUP;

/// \brief <em>Identifies the image list that is used for footer items</em>
///
/// This is an undocumented value for the \c wParam parameter of the \c LVM_GETIMAGELIST and
/// \c LVM_SETIMAGELIST messages.
///
/// \remarks Requires comctl32.dll version 6.10 or higher.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/bb774943.aspx">LVM_GETIMAGELIST</a>,
///     <a href="https://msdn.microsoft.com/en-us/library/bb761178.aspx">LVM_SETIMAGELIST</a>
#define LVSIL_FOOTERITEMS 4

#ifdef _DEBUG
	/// \brief <em>Bottom-aligns items in 'Icons' and 'Small Icons' view</em>
	///
	/// This is an undocumented window style that causes the listview to bottom-align items in 'Icons' and
	/// 'Small Icons' view.
	///
	/// \sa LVS_ALIGNRIGHT
	#define LVS_ALIGNBOTTOM 0x0400
	/// \brief <em>Right-aligns items in 'Icons' and 'Small Icons' view</em>
	///
	/// This is an undocumented window style that causes the listview to right-align items in 'Icons' and
	/// 'Small Icons' view.
	///
	/// \sa LVS_ALIGNBOTTOM
	#define LVS_ALIGNRIGHT 0x0C00

	#define LVBKIF_STYLE_STRETCH 0x00000000     // TODO: Find out the correct value
#endif


#ifndef DOXYGEN_SHOULD_SKIP_THIS
	// missing in commctrl.h
	#define LVS_EX_DRAWIMAGEASYNC 0x04000000
	#define LVN_ASYNCDRAWN (LVN_FIRST - 86)
	#define LVADPART_ITEM 0
	#define LVADPART_GROUP 1
	typedef struct tagNMLVASYNCDRAWN
	{
		NMHDR hdr;
		IMAGELISTDRAWPARAMS *pimldp;
		HRESULT hr;
		int iPart;
		int iItem;
		int iSubItem;
		LPARAM lParam;
		// out params
		DWORD dwRetFlags;
		int iRetImageIndex;
	} NMLVASYNCDRAWN;
#endif