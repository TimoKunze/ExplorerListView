//////////////////////////////////////////////////////////////////////
/// \class Proxy_IExplorerListViewEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IExplorerListViewEvents interface</em>
///
/// \if UNICODE
///   \sa ExplorerListView, ExLVwLibU::_IExplorerListViewEvents
/// \else
///   \sa ExplorerListView, ExLVwLibA::_IExplorerListViewEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IExplorerListViewEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IExplorerListViewEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c AbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::AbortedDrag, ExplorerListView::Raise_AbortedDrag,
	///       Fire_Drop, Fire_HeaderAbortedDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::AbortedDrag, ExplorerListView::Raise_AbortedDrag,
	///       Fire_Drop, Fire_HeaderAbortedDrag
	/// \endif
	HRESULT Fire_AbortedDrag(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ABORTEDDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c AfterScroll event</em>
	///
	/// \param[in] dx The number of steps that the control was scrolled horizontally.
	/// \param[in] dy The number of steps that the control was scrolled vertically.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::AfterScroll, ExplorerListView::Raise_AfterScroll,
	///       Fire_BeforeScroll
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::AfterScroll, ExplorerListView::Raise_AfterScroll,
	///       Fire_BeforeScroll
	/// \endif
	HRESULT Fire_AfterScroll(LONG dx, LONG dy)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = dx;
				p[0] = dy;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_AFTERSCROLL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BeforeScroll event</em>
	///
	/// \param[in] dx The number of steps that the control is about to be scroll horizontally.
	/// \param[in] dy The number of steps that the control is about to be scroll vertically.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::BeforeScroll, ExplorerListView::Raise_BeforeScroll,
	///       Fire_AfterScroll
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::BeforeScroll, ExplorerListView::Raise_BeforeScroll,
	///       Fire_AfterScroll
	/// \endif
	HRESULT Fire_BeforeScroll(LONG dx, LONG dy)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = dx;
				p[0] = dy;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_BEFORESCROLL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BeginColumnResizing event</em>
	///
	/// \param[in] pColumn The column being resized.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort column resizing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::BeginColumnResizing,
	///       ExplorerListView::Raise_BeginColumnResizing, Fire_ResizingColumn, Fire_EndColumnResizing
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::BeginColumnResizing,
	///       ExplorerListView::Raise_BeginColumnResizing, Fire_ResizingColumn, Fire_EndColumnResizing
	/// \endif
	HRESULT Fire_BeginColumnResizing(IListViewColumn* pColumn, VARIANT_BOOL* pCancel)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pColumn;
				p[0].pboolVal = pCancel;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_BEGINCOLUMNRESIZING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BeginMarqueeSelection event</em>
	///
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort marquee selection; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::BeginMarqueeSelection,
	///       ExplorerListView::Raise_BeginMarqueeSelection, Fire_ItemSelectionChanged,
	///       Fire_SelectedItemRange
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::BeginMarqueeSelection,
	///       ExplorerListView::Raise_BeginMarqueeSelection, Fire_ItemSelectionChanged,
	///       Fire_SelectedItemRange
	/// \endif
	HRESULT Fire_BeginMarqueeSelection(VARIANT_BOOL* pCancel)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].pboolVal = pCancel;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_BEGINMARQUEESELECTION, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CacheItemsHint event</em>
	///
	/// \param[in] pFirstItem The first item to cache.
	/// \param[in] pLastItem The last item to cache.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CacheItemsHint, ExplorerListView::Raise_CacheItemsHint,
	///       Fire_ItemGetDisplayInfo
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CacheItemsHint, ExplorerListView::Raise_CacheItemsHint,
	///       Fire_ItemGetDisplayInfo
	/// \endif
	HRESULT Fire_CacheItemsHint(IListViewItem* pFirstItem, IListViewItem* pLastItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pFirstItem;
				p[0] = pLastItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CACHEITEMSHINT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CancelSubItemEdit event</em>
	///
	/// \param[in] pListSubItem The sub-item that has been edited.
	/// \param[in] editMode Specifies how the label-edit mode has been entered. Any of the values defined by
	///            the \c SubItemEditModeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CancelSubItemEdit,
	///       ExplorerListView::Raise_CancelSubItemEdit, Fire_GetSubItemControl,
	///       Fire_ConfigureSubItemControl, Fire_EndSubItemEdit
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CancelSubItemEdit,
	///       ExplorerListView::Raise_CancelSubItemEdit, Fire_GetSubItemControl,
	///       Fire_ConfigureSubItemControl, Fire_EndSubItemEdit
	/// \endif
	HRESULT Fire_CancelSubItemEdit(IListViewSubItem* pListSubItem, SubItemEditModeConstants editMode)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListSubItem;
				p[0].lVal = static_cast<LONG>(editMode);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CANCELSUBITEMEDIT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CaretChanged event</em>
	///
	/// \param[in] pPreviousCaretItem The previous caret item.
	/// \param[in] pNewCaretItem The new caret item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CaretChanged, ExplorerListView::Raise_CaretChanged,
	///       Fire_ItemSelectionChanged
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CaretChanged, ExplorerListView::Raise_CaretChanged,
	///       Fire_ItemSelectionChanged
	/// \endif
	HRESULT Fire_CaretChanged(IListViewItem* pPreviousCaretItem, IListViewItem* pNewCaretItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousCaretItem;
				p[0] = pNewCaretItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CARETCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedSortOrder event</em>
	///
	/// \param[in] previousSortOrder The control's old sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in] newSortOrder The control's new sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ChangedSortOrder,
	///       ExplorerListView::Raise_ChangedSortOrder, Fire_ChangingSortOrder
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ChangedSortOrder,
	///       ExplorerListView::Raise_ChangedSortOrder, Fire_ChangingSortOrder
	/// \endif
	HRESULT Fire_ChangedSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].lVal = static_cast<LONG>(previousSortOrder);		p[1].vt = VT_I4;
				p[0].lVal = static_cast<LONG>(newSortOrder);				p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CHANGEDSORTORDER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangedWorkAreas event</em>
	///
	/// \param[in] pWorkAreas A collection of the new working areas.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ChangedWorkAreas,
	///       ExplorerListView::Raise_ChangedWorkAreas, Fire_ChangingWorkAreas
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ChangedWorkAreas,
	///       ExplorerListView::Raise_ChangedWorkAreas, Fire_ChangingWorkAreas
	/// \endif
	HRESULT Fire_ChangedWorkAreas(IListViewWorkAreas* pWorkAreas)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pWorkAreas;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CHANGEDWORKAREAS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangingSortOrder event</em>
	///
	/// \param[in] previousSortOrder The control's old sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in] newSortOrder The control's new sort order. Any of the values defined by the
	///            \c SortOrderConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort redefining; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ChangingSortOrder,
	///       ExplorerListView::Raise_ChangingSortOrder, Fire_ChangedSortOrder
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ChangingSortOrder,
	///       ExplorerListView::Raise_ChangingSortOrder, Fire_ChangedSortOrder
	/// \endif
	HRESULT Fire_ChangingSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2].lVal = static_cast<LONG>(previousSortOrder);		p[2].vt = VT_I4;
				p[1].lVal = static_cast<LONG>(newSortOrder);				p[1].vt = VT_I4;
				p[0].pboolVal = pCancelChange;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CHANGINGSORTORDER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChangingWorkAreas event</em>
	///
	/// \param[in] pWorkAreas A collection of the new working areas.
	/// \param[in,out] pCancelChanges If \c VARIANT_TRUE, the caller should abort redefining; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ChangingWorkAreas,
	///       ExplorerListView::Raise_ChangingWorkAreas, Fire_ChangedWorkAreas
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ChangingWorkAreas,
	///       ExplorerListView::Raise_ChangingWorkAreas, Fire_ChangedWorkAreas
	/// \endif
	HRESULT Fire_ChangingWorkAreas(IVirtualListViewWorkAreas* pWorkAreas, VARIANT_BOOL* pCancelChanges)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pWorkAreas;
				p[0].pboolVal = pCancelChanges;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CHANGINGWORKAREAS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
	/// \param[in] pListSubItem The clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::Click, ExplorerListView::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::Click, ExplorerListView::Raise_Click, Fire_DblClick,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_Click(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CollapsedGroup event</em>
	///
	/// \param[in] pGroup The group that was expanded.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CollapsedGroup, ExplorerListView::Raise_CollapsedGroup,
	///       Fire_ExpandedGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CollapsedGroup, ExplorerListView::Raise_CollapsedGroup,
	///       Fire_ExpandedGroup
	/// \endif
	HRESULT Fire_CollapsedGroup(IListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLLAPSEDGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnBeginDrag event</em>
	///
	/// \param[in] pColumn The column header that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HeaderHitTestConstants enumeration are valid.
	/// \param[in,out] pDoAutomaticDragDrop If set to \c VARIANT_TRUE, the header control should handle
	///                column drag'n'drop itself; otherwise the client is responsible.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnBeginDrag, ExplorerListView::Raise_ColumnBeginDrag,
	///       Fire_ColumnEndAutoDragDrop, Fire_ItemBeginDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnBeginDrag, ExplorerListView::Raise_ColumnBeginDrag,
	///       Fire_ColumnEndAutoDragDrop, Fire_ItemBeginDrag
	/// \endif
	HRESULT Fire_ColumnBeginDrag(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails, VARIANT_BOOL* pDoAutomaticDragDrop)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pColumn;
				p[5] = button;																			p[5].vt = VT_I2;
				p[4] = shift;																				p[4].vt = VT_I2;
				p[3] = x;																						p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																						p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);			p[1].vt = VT_I4;
				p[0].pboolVal = pDoAutomaticDragDrop;								p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnClick event</em>
	///
	/// \param[in] pColumn The clicked column header.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was clicked. Some of the values defined
	///            by the \c HeaderHitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnClick, ExplorerListView::Raise_ColumnClick,
	///       Fire_HeaderClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnClick, ExplorerListView::Raise_ColumnClick,
	///       Fire_HeaderClick
	/// \endif
	HRESULT Fire_ColumnClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnDropDown event</em>
	///
	/// \param[in] pColumn The column header whose drop-down menu should be displayed.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the drop-down menu's proposed position relative to the
	///            header control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the drop-down menu's proposed position relative to the
	///            header control's upper-left corner.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the \c ShellListView
	///                control from showing the default drop-down menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnDropDown, ExplorerListView::Raise_ColumnDropDown,
	///       Fire_ColumnClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnDropDown, ExplorerListView::Raise_ColumnDropDown,
	///       Fire_ColumnClick
	/// \endif
	HRESULT Fire_ColumnDropDown(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;											p[4].vt = VT_I2;
				p[3] = shift;												p[3].vt = VT_I2;
				p[2] = x;														p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;														p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pShowDefaultMenu;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNDROPDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnEndAutoDragDrop event</em>
	///
	/// \param[in] pColumn The column header that was dragged.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HeaderHitTestConstants enumeration are valid.
	/// \param[in,out] pDoAutomaticDrop If set to \c VARIANT_TRUE, the header control should handle the
	///                column header drop itself; otherwise the client is responsible.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnEndAutoDragDrop,
	///       ExplorerListView::Raise_ColumnEndAutoDragDrop, Fire_ColumnBeginDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnEndAutoDragDrop,
	///       ExplorerListView::Raise_ColumnEndAutoDragDrop, Fire_ColumnBeginDrag
	/// \endif
	HRESULT Fire_ColumnEndAutoDragDrop(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails, VARIANT_BOOL* pDoAutomaticDrop)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pColumn;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pDoAutomaticDrop;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNENDAUTODRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnMouseEnter event</em>
	///
	/// \param[in] pColumn The column header that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HeaderHitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnMouseEnter, ExplorerListView::Raise_ColumnMouseEnter,
	///       Fire_ColumnMouseLeave, Fire_HeaderMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnMouseEnter, ExplorerListView::Raise_ColumnMouseEnter,
	///       Fire_ColumnMouseLeave, Fire_HeaderMouseMove
	/// \endif
	HRESULT Fire_ColumnMouseEnter(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnMouseLeave event</em>
	///
	/// \param[in] pColumn The column header that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HeaderHitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnMouseLeave, ExplorerListView::Raise_ColumnMouseLeave,
	///       Fire_ColumnMouseEnter, Fire_HeaderMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnMouseLeave, ExplorerListView::Raise_ColumnMouseLeave,
	///       Fire_ColumnMouseEnter, Fire_HeaderMouseMove
	/// \endif
	HRESULT Fire_ColumnMouseLeave(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnStateImageChanged event</em>
	///
	/// \param[in] pColumn The column whose state image was changed.
	/// \param[in] previousStateImageIndex The one-based index of the column's previous state image.
	/// \param[in] newStateImageIndex The one-based index of the column's new state image.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnStateImageChanged,
	///       ExplorerListView::Raise_ColumnStateImageChanged, Fire_ColumnStateImageChanging,
	///       Fire_ItemStateImageChanged
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnStateImageChanged,
	///       ExplorerListView::Raise_ColumnStateImageChanged, Fire_ColumnStateImageChanging,
	///       Fire_ItemStateImageChanged
	/// \endif
	HRESULT Fire_ColumnStateImageChanged(IListViewColumn* pColumn, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pColumn;
				p[2] = previousStateImageIndex;
				p[1] = newStateImageIndex;
				p[0] = static_cast<LONG>(causedBy);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNSTATEIMAGECHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ColumnStateImageChanging event</em>
	///
	/// \param[in] pColumn The column whose state image shall be changed.
	/// \param[in] previousStateImageIndex The one-based index of the column's previous state image.
	/// \param[in,out] pNewStateImageIndex The one-based index of the column's new state image. The client
	///                may change this value.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should prevent the state image from
	///                changing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ColumnStateImageChanging,
	///       ExplorerListView::Raise_ColumnStateImageChanging, Fire_ColumnStateImageChanged,
	///       Fire_ItemStateImageChanging
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ColumnStateImageChanging,
	///       ExplorerListView::Raise_ColumnStateImageChanging, Fire_ColumnStateImageChanged,
	///       Fire_ItemStateImageChanging
	/// \endif
	HRESULT Fire_ColumnStateImageChanging(IListViewColumn* pColumn, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pColumn;
				p[3] = previousStateImageIndex;
				p[2].plVal = pNewStateImageIndex;				p[2].vt = VT_I4 | VT_BYREF;
				p[1] = static_cast<LONG>(causedBy);			p[1].vt = VT_I4;
				p[0].pboolVal = pCancelChange;					p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COLUMNSTATEIMAGECHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CompareGroups event</em>
	///
	/// \param[in] pFirstGroup The first group to compare.
	/// \param[in] pSecondGroup The second group to compare.
	/// \param[out] pResult Receives one of the values defined by the \c CompareResultConstants enumeration.
	///             If \c GroupSortOrder is set to \c soDescending, the caller should invert this value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CompareGroups, ExplorerListView::Raise_CompareGroups
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CompareGroups, ExplorerListView::Raise_CompareGroups
	/// \endif
	HRESULT Fire_CompareGroups(IListViewGroup* pFirstGroup, IListViewGroup* pSecondGroup, CompareResultConstants* pResult)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pFirstGroup;
				p[1] = pSecondGroup;
				p[0].plVal = reinterpret_cast<PLONG>(pResult);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COMPAREGROUPS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CompareItems event</em>
	///
	/// \param[in] pFirstItem The first item to compare.
	/// \param[in] pSecondItem The second item to compare.
	/// \param[out] pResult Receives one of the values defined by the \c CompareResultConstants enumeration.
	///             If \c SortOrder is set to \c soDescending, the caller should invert this value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CompareItems, ExplorerListView::Raise_CompareItems
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CompareItems, ExplorerListView::Raise_CompareItems
	/// \endif
	HRESULT Fire_CompareItems(IListViewItem* pFirstItem, IListViewItem* pSecondItem, CompareResultConstants* pResult)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pFirstItem;
				p[1] = pSecondItem;
				p[0].plVal = reinterpret_cast<PLONG>(pResult);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_COMPAREITEMS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ConfigureSubItemControl event</em>
	///
	/// \param[in] pListSubItem The sub-item that the representation control needs to be configured for.
	/// \param[in] controlKind The kind of representation control being configured. Any of the values
	///            defined by the \c SubItemControlKindConstants enumeration are valid.
	/// \param[in] editMode Specifies how the label-edit mode is being entered. Any of the values defined by
	///            the \c SubItemEditModeConstants enumeration is valid.
	/// \param[in] subItemControl The representation control that needs to be configured. Any of the
	///            values defined by the \c SubItemControlConstants enumeration are valid.
	/// \param[in,out] pThemeAppName Specifies the application name of the theme to apply when drawing
	///                the sub-item control. For instance this value can be set to "explorer" to make
	///                the sub-item control be drawn like in Windows Explorer.
	/// \param[in,out] pThemeIDList Specifies a semicolon-separated list of CLSID names to use in place
	///                of the names specified by the window's class. This value is used to refine the
	///                search for a visual style to apply. For instance there might be different visual
	///                styles available for different usages of the same window class.
	/// \param[in,out] phFont Specifies the font to apply to the sub-item control.
	/// \param[in,out] pTextColor Specifies the text color to apply to the sub-item control.
	/// \param[in,out] ppPropertyDescription An object that implements the \c IPropertyDescription interface.
	///                This object is used for a more detailed configuration of the sub-item control.
	///                Some built-in sub-item controls like the \c sicDropList control won't work without
	///                specifying an \c IPropertyDescription implementation.
	/// \param[in,out] pPropertyValue Specifies the address of a \c PROPVARIANT structure that holds the
	///                sub-item's current value. Sub-items can be thought of as representing various
	///                properties of the item that they belong to. These properties can be of any type,
	///                not only strings. The \c PROPVARIANT type is similar to Visual Basic's \c Variant
	///                type and can hold any other type, for instance integer numbers, floating-point
	///                numbers and objects.\n
	///                Use the <a href="https://msdn.microsoft.com/en-us/library/bb762286.aspx">PROPVARIANT
	///                and VARIANT API functions</a> to work with the \c PROPVARIANT data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ConfigureSubItemControl,
	///       ExplorerListView::Raise_ConfigureSubItemControl, Fire_GetSubItemControl, Fire_EndSubItemEdit,
	///       Fire_CancelSubItemEdit, Fire_CustomDraw
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ConfigureSubItemControl,
	///       ExplorerListView::Raise_ConfigureSubItemControl, Fire_GetSubItemControl, Fire_EndSubItemEdit,
	///       Fire_CancelSubItemEdit, Fire_CustomDraw
	/// \endif
	HRESULT Fire_ConfigureSubItemControl(IListViewSubItem* pListSubItem, SubItemControlKindConstants controlKind, SubItemEditModeConstants editMode, SubItemControlConstants subItemControl, BSTR* pThemeAppName, BSTR* pThemeIDList, LONG* phFont, OLE_COLOR* pTextColor, IUnknown** ppPropertyDescription, LONG pPropertyValue)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pListSubItem;
				p[8] = controlKind;
				p[7].lVal = static_cast<LONG>(editMode);						p[7].vt = VT_I4;
				p[6] = subItemControl;
				p[5].pbstrVal = pThemeAppName;											p[5].vt = VT_BSTR | VT_BYREF;
				p[4].pbstrVal = pThemeIDList;												p[4].vt = VT_BSTR | VT_BYREF;
				p[3].plVal = phFont;																p[3].vt = VT_I4 | VT_BYREF;
				p[2].plVal = reinterpret_cast<LONG*>(pTextColor);		p[2].vt = VT_I4 | VT_BYREF;
				LONG pPropertyDescription = *reinterpret_cast<LONG*>(ppPropertyDescription);
				p[1].plVal = &pPropertyDescription;									p[1].vt = VT_I4 | VT_BYREF;
				p[0] = pPropertyValue;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CONFIGURESUBITEMCONTROL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(*ppPropertyDescription && pPropertyDescription != *reinterpret_cast<LONG*>(ppPropertyDescription)) {
					(*ppPropertyDescription)->Release();
					*ppPropertyDescription = NULL;
				}
				*ppPropertyDescription = *reinterpret_cast<IUnknown**>(&pPropertyDescription);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] pListItem The item that the context menu refers to. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the context menu refers to. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the \c ShellListView
	///                control from showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ContextMenu, ExplorerListView::Raise_ContextMenu,
	///       Fire_RClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ContextMenu, ExplorerListView::Raise_ContextMenu,
	///       Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pListItem;
				p[6] = pListSubItem;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pShowDefaultMenu;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CreatedEditControlWindow,
	///       ExplorerListView::Raise_CreatedEditControlWindow, Fire_DestroyedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CreatedEditControlWindow,
	///       ExplorerListView::Raise_CreatedEditControlWindow, Fire_DestroyedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \endif
	HRESULT Fire_CreatedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CREATEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CreatedHeaderControlWindow event</em>
	///
	/// \param[in] hWndHeader The header control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CreatedHeaderControlWindow,
	///       ExplorerListView::Raise_CreatedHeaderControlWindow, Fire_DestroyedHeaderControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CreatedHeaderControlWindow,
	///       ExplorerListView::Raise_CreatedHeaderControlWindow, Fire_DestroyedHeaderControlWindow
	/// \endif
	HRESULT Fire_CreatedHeaderControlWindow(LONG hWndHeader)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndHeader;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CREATEDHEADERCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CustomDraw event</em>
	///
	/// \param[in] pListItem The item that the notification refers to. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the notification refers to. May be \c NULL.
	/// \param[in] drawAllItems If \c VARIANT_TRUE, all items are to be drawn.
	/// \param[in,out] pTextColor An \c OLE_COLOR value specifying the color to draw the item's text in.
	///                The client may change this value.
	/// \param[in,out] pTextBackColor An \c OLE_COLOR value specifying the color to fill the item's
	///                text background with. The client may change this value.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Any of the values
	///            defined by the \c CustomDrawStageConstants enumeration is valid.
	/// \param[in] itemState The item's current state (focused, selected etc.). Most of the values
	///            defined by the \c CustomDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::CustomDraw, ExplorerListView::Raise_CustomDraw,
	///       Fire_GroupCustomDraw, Fire_HeaderCustomDraw
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::CustomDraw, ExplorerListView::Raise_CustomDraw,
	///       Fire_GroupCustomDraw, Fire_HeaderCustomDraw
	/// \endif
	HRESULT Fire_CustomDraw(IListViewItem* pListItem, IListViewSubItem* pListSubItem, VARIANT_BOOL drawAllItems, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pListItem;
				p[8] = pListSubItem;
				p[7] = drawAllItems;
				p[6].plVal = reinterpret_cast<PLONG>(pTextColor);						p[6].vt = VT_I4 | VT_BYREF;
				p[5].plVal = reinterpret_cast<PLONG>(pTextBackColor);				p[5].vt = VT_I4 | VT_BYREF;
				p[4].lVal = static_cast<LONG>(drawStage);										p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(itemState);										p[3].vt = VT_I4;
				p[2] = hDC;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4C7D8A62-0F3A-41a7-BD58-FF25497D8451}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{00B2ADA1-29FD-45ce-A4C3-A2B85928FADD}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo;
				p[1].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_CUSTOMDRAW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[1].pvRecord);
				}
			}
		}

		/* Although pTextColor and pTextBackColor are of type OLE_COLOR, they may contain RGB colors only.
		   So convert OLE_COLOR colors into RGB colors. */
		if(*pTextColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pTextColor, NULL, &color);
			*pTextColor = static_cast<OLE_COLOR>(color);
		}
		if(*pTextBackColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pTextBackColor, NULL, &color);
			*pTextBackColor = static_cast<OLE_COLOR>(color);
		}

		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
	/// \param[in] pListSubItem The double-clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::DblClick, ExplorerListView::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::DblClick, ExplorerListView::Raise_DblClick, Fire_Click,
	///       Fire_MDblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::DestroyedControlWindow,
	///       ExplorerListView::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::DestroyedControlWindow,
	///       ExplorerListView::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \endif
	HRESULT Fire_DestroyedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedEditControlWindow event</em>
	///
	/// \param[in] hWndEdit The edit control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::DestroyedEditControlWindow,
	///       ExplorerListView::Raise_DestroyedEditControlWindow, Fire_CreatedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::DestroyedEditControlWindow,
	///       ExplorerListView::Raise_DestroyedEditControlWindow, Fire_CreatedEditControlWindow,
	///       Fire_StartingLabelEditing
	/// \endif
	HRESULT Fire_DestroyedEditControlWindow(LONG hWndEdit)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndEdit;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_DESTROYEDEDITCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedHeaderControlWindow event</em>
	///
	/// \param[in] hWndHeader The header control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::DestroyedHeaderControlWindow,
	///       ExplorerListView::Raise_DestroyedHeaderControlWindow, Fire_CreatedHeaderControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::DestroyedHeaderControlWindow,
	///       ExplorerListView::Raise_DestroyedHeaderControlWindow, Fire_CreatedHeaderControlWindow
	/// \endif
	HRESULT Fire_DestroyedHeaderControlWindow(LONG hWndHeader)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndHeader;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_DESTROYEDHEADERCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DragMouseMove event</em>
	///
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::DragMouseMove, ExplorerListView::Raise_DragMouseMove,
	///       Fire_MouseMove, Fire_OLEDragMouseMove, Fire_HeaderDragMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::DragMouseMove, ExplorerListView::Raise_DragMouseMove,
	///       Fire_MouseMove, Fire_OLEDragMouseMove, Fire_HeaderDragMouseMove
	/// \endif
	HRESULT Fire_DragMouseMove(IListViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[7].vt = VT_DISPATCH | VT_BYREF;
				p[6] = button;																									p[6].vt = VT_I2;
				p[5] = shift;																										p[5].vt = VT_I2;
				p[4] = x;																												p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																												p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_DRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Drop event</em>
	///
	/// \param[in] pDropTarget The item object that is the nearest one from the mouse cursor's position.
	///            If the mouse cursor's position lies outside the control's client area, this parameter
	///            should be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::Drop, ExplorerListView::Raise_Drop, Fire_AbortedDrag,
	///       Fire_HeaderDrop
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::Drop, ExplorerListView::Raise_Drop, Fire_AbortedDrag,
	///       Fire_HeaderDrop
	/// \endif
	HRESULT Fire_Drop(IListViewItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_DROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditChange event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditChange, ExplorerListView::Raise_EditChange,
	///       Fire_EditKeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditChange, ExplorerListView::Raise_EditChange,
	///       Fire_EditKeyPress
	/// \endif
	HRESULT Fire_EditChange(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITCHANGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditClick, ExplorerListView::Raise_EditClick,
	///       Fire_EditDblClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditClick, ExplorerListView::Raise_EditClick,
	///       Fire_EditDblClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the contained edit
	///                control from showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditContextMenu,
	///       ExplorerListView::Raise_EditContextMenu, Fire_EditRClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditContextMenu,
	///       ExplorerListView::Raise_EditContextMenu, Fire_EditRClick
	/// \endif
	HRESULT Fire_EditContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = button;											p[4].vt = VT_I2;
				p[3] = shift;												p[3].vt = VT_I2;
				p[2] = x;														p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;														p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pShowDefaultMenu;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditDblClick, ExplorerListView::Raise_EditDblClick,
	///       Fire_EditClick, Fire_EditMDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditDblClick, ExplorerListView::Raise_EditDblClick,
	///       Fire_EditClick, Fire_EditMDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \endif
	HRESULT Fire_EditDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditGotFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditGotFocus, ExplorerListView::Raise_EditGotFocus,
	///       Fire_EditLostFocus
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditGotFocus, ExplorerListView::Raise_EditGotFocus,
	///       Fire_EditLostFocus
	/// \endif
	HRESULT Fire_EditGotFocus(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITGOTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditKeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditKeyDown, ExplorerListView::Raise_EditKeyDown,
	///       Fire_EditKeyUp, Fire_EditKeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditKeyDown, ExplorerListView::Raise_EditKeyDown,
	///       Fire_EditKeyUp, Fire_EditKeyPress
	/// \endif
	HRESULT Fire_EditKeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITKEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditKeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditKeyPress, ExplorerListView::Raise_EditKeyPress,
	///       Fire_EditKeyDown, Fire_EditKeyUp
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditKeyPress, ExplorerListView::Raise_EditKeyPress,
	///       Fire_EditKeyDown, Fire_EditKeyUp
	/// \endif
	HRESULT Fire_EditKeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITKEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditKeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditKeyUp, ExplorerListView::Raise_EditKeyUp,
	///       Fire_EditKeyDown, Fire_EditKeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditKeyUp, ExplorerListView::Raise_EditKeyUp,
	///       Fire_EditKeyDown, Fire_EditKeyPress
	/// \endif
	HRESULT Fire_EditKeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITKEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditLostFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditLostFocus, ExplorerListView::Raise_EditLostFocus,
	///       Fire_EditGotFocus
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditLostFocus, ExplorerListView::Raise_EditLostFocus,
	///       Fire_EditGotFocus
	/// \endif
	HRESULT Fire_EditLostFocus(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITLOSTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMClick, ExplorerListView::Raise_EditMClick,
	///       Fire_EditMDblClick, Fire_EditClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMClick, ExplorerListView::Raise_EditMClick,
	///       Fire_EditMDblClick, Fire_EditClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMDblClick, ExplorerListView::Raise_EditMDblClick,
	///       Fire_EditMClick, Fire_EditDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMDblClick, ExplorerListView::Raise_EditMDblClick,
	///       Fire_EditMClick, Fire_EditDblClick, Fire_EditRDblClick, Fire_EditXDblClick
	/// \endif
	HRESULT Fire_EditMDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseDown, ExplorerListView::Raise_EditMouseDown,
	///       Fire_EditMouseUp, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseDown, ExplorerListView::Raise_EditMouseDown,
	///       Fire_EditMouseUp, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseEnter, ExplorerListView::Raise_EditMouseEnter,
	///       Fire_EditMouseLeave, Fire_EditMouseHover, Fire_EditMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseEnter, ExplorerListView::Raise_EditMouseEnter,
	///       Fire_EditMouseLeave, Fire_EditMouseHover, Fire_EditMouseMove
	/// \endif
	HRESULT Fire_EditMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseHover, ExplorerListView::Raise_EditMouseHover,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseHover, ExplorerListView::Raise_EditMouseHover,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseMove
	/// \endif
	HRESULT Fire_EditMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseLeave, ExplorerListView::Raise_EditMouseLeave,
	///       Fire_EditMouseEnter, Fire_EditMouseHover, Fire_EditMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseLeave, ExplorerListView::Raise_EditMouseLeave,
	///       Fire_EditMouseEnter, Fire_EditMouseHover, Fire_EditMouseMove
	/// \endif
	HRESULT Fire_EditMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseMove, ExplorerListView::Raise_EditMouseMove,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseDown, Fire_EditMouseUp,
	///       Fire_EditMouseWheel
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseMove, ExplorerListView::Raise_EditMouseMove,
	///       Fire_EditMouseEnter, Fire_EditMouseLeave, Fire_EditMouseDown, Fire_EditMouseUp,
	///       Fire_EditMouseWheel
	/// \endif
	HRESULT Fire_EditMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseUp event</em>
	///
	/// \param[in] button The released mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseUp, ExplorerListView::Raise_EditMouseUp,
	///       Fire_EditMouseDown, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseUp, ExplorerListView::Raise_EditMouseUp,
	///       Fire_EditMouseDown, Fire_EditClick, Fire_EditMClick, Fire_EditRClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditMouseWheel event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditMouseWheel, ExplorerListView::Raise_EditMouseWheel,
	///       Fire_EditMouseMove, Fire_MouseWheel, Fire_HeaderMouseWheel
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditMouseWheel, ExplorerListView::Raise_EditMouseWheel,
	///       Fire_EditMouseMove, Fire_MouseWheel, Fire_HeaderMouseWheel
	/// \endif
	HRESULT Fire_EditMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = button;																p[5].vt = VT_I2;
				p[4] = shift;																	p[4].vt = VT_I2;
				p[3] = x;																			p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																			p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(scrollAxis);		p[1].vt = VT_I4;
				p[0] = wheelDelta;														p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITMOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditRClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditRClick, ExplorerListView::Raise_EditRClick,
	///       Fire_EditContextMenu, Fire_EditRDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditRClick, ExplorerListView::Raise_EditRClick,
	///       Fire_EditContextMenu, Fire_EditRDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditXClick
	/// \endif
	HRESULT Fire_EditRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITRCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditRDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditRDblClick, ExplorerListView::Raise_EditRDblClick,
	///       Fire_EditRClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditXDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditRDblClick, ExplorerListView::Raise_EditRDblClick,
	///       Fire_EditRClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditXDblClick
	/// \endif
	HRESULT Fire_EditRDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITRDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditXClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the contained edit
	///            control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditXClick, ExplorerListView::Raise_EditXClick,
	///       Fire_EditXDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditRClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditXClick, ExplorerListView::Raise_EditXClick,
	///       Fire_EditXDblClick, Fire_EditClick, Fire_EditMClick, Fire_EditRClick
	/// \endif
	HRESULT Fire_EditXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITXCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EditXDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the contained
	///            edit control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EditXDblClick, ExplorerListView::Raise_EditXDblClick,
	///       Fire_EditXClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditRDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EditXDblClick, ExplorerListView::Raise_EditXDblClick,
	///       Fire_EditXClick, Fire_EditDblClick, Fire_EditMDblClick, Fire_EditRDblClick
	/// \endif
	HRESULT Fire_EditXDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EDITXDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EmptyMarkupTextLinkClick event</em>
	///
	/// \param[in] linkIndex The zero-based index of the link that was clicked.
	/// \param[in] button The mouse button that was pressed during the click. Any of the values defined
	///            by VB's \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined by the
	///            \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EmptyMarkupTextLinkClick,
	///       ExplorerListView::Raise_EmptyMarkupTextLinkClick, Fire_Click
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EmptyMarkupTextLinkClick,
	///       ExplorerListView::Raise_EmptyMarkupTextLinkClick, Fire_Click
	/// \endif
	HRESULT Fire_EmptyMarkupTextLinkClick(LONG linkIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = linkIndex;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EMPTYMARKUPTEXTLINKCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EndColumnResizing event</em>
	///
	/// \param[in] pColumn The column that was resized.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EndColumnResizing,
	///       ExplorerListView::Raise_EndColumnResizing, Fire_BeginColumnResizing, Fire_ResizingColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EndColumnResizing,
	///       ExplorerListView::Raise_EndColumnResizing, Fire_BeginColumnResizing, Fire_ResizingColumn
	/// \endif
	HRESULT Fire_EndColumnResizing(IListViewColumn* pColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pColumn;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ENDCOLUMNRESIZING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EndSubItemEdit event</em>
	///
	/// \param[in] pListSubItem The sub-item that has been edited.
	/// \param[in] editMode Specifies how the label-edit mode has been entered. Any of the values defined by
	///            the \c SubItemEditModeConstants enumeration is valid.
	/// \param[in] pPropertyKey Specifies the address of a \c PROPERTYKEY structure that identifies the
	///            property that the secified sub-item is representing. This is the \c PROPERTYKEY
	///            structure that has been retrieved by the \c GetPropertyKey method of the
	///            \c IPropertyDescription object that previously has been provided by the
	///            \c ConfigureSubItemControl event handler.
	/// \param[in] pPropertyValue Specifies the address of a \c PROPVARIANT structure that holds the
	///            sub-item's new value. Sub-items can be thought of as representing various properties
	///            of the item that they belong to. These properties can be of any type, not only
	///            strings. The \c PROPVARIANT type is similar to Visual Basic's \c Variant type and can
	///            hold any other type, for instance integer numbers, floating-point numbers and
	///            objects.\n
	///            Use the <a href="https://msdn.microsoft.com/en-us/library/bb762286.aspx">PROPVARIANT
	///            and VARIANT API functions</a> to work with the \c PROPVARIANT data.
	/// \param[in,out] pCancel If set to \c VARIANT_TRUE, the editing of the sub-item should not be processed
	///                any further.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::EndSubItemEdit, ExplorerListView::Raise_EndSubItemEdit,
	///       Fire_GetSubItemControl, Fire_ConfigureSubItemControl, Fire_CancelSubItemEdit
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::EndSubItemEdit, ExplorerListView::Raise_EndSubItemEdit,
	///       Fire_GetSubItemControl, Fire_ConfigureSubItemControl, Fire_CancelSubItemEdit
	/// \endif
	HRESULT Fire_EndSubItemEdit(IListViewSubItem* pListSubItem, SubItemEditModeConstants editMode, LONG pPropertyKey, LONG pPropertyValue, VARIANT_BOOL* pCancel)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pListSubItem;
				p[3].lVal = static_cast<LONG>(editMode);		p[3].vt = VT_I4;
				p[2] = pPropertyKey;
				p[1] = pPropertyValue;
				p[0].pboolVal = pCancel;										p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ENDSUBITEMEDIT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ExpandedGroup event</em>
	///
	/// \param[in] pGroup The group that was expanded.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ExpandedGroup, ExplorerListView::Raise_ExpandedGroup,
	///       Fire_CollapsedGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ExpandedGroup, ExplorerListView::Raise_ExpandedGroup,
	///       Fire_CollapsedGroup
	/// \endif
	HRESULT Fire_ExpandedGroup(IListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_EXPANDEDGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FilterButtonClick event</em>
	///
	/// \param[in] pColumn The column header whose 'Apply Filter' button was clicked.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] pFilterButtonRectangle The bounding rectangle of the clicked button's client area
	///            (in pixels).
	/// \param[in,out] pRaiseFilterChanged If \c VARIANT_TRUE, the caller should raise the \c FilterChanged
	///                event; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FilterButtonClick,
	///       ExplorerListView::Raise_FilterButtonClick, Fire_FilterChanged, Fire_HeaderClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FilterButtonClick,
	///       ExplorerListView::Raise_FilterButtonClick, Fire_FilterChanged, Fire_HeaderClick
	/// \endif
	HRESULT Fire_FilterButtonClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, RECTANGLE* pFilterButtonRectangle, VARIANT_BOOL* pRaiseFilterChanged)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pColumn;
				p[5] = button;													p[5].vt = VT_I2;
				p[4] = shift;														p[4].vt = VT_I2;
				p[3] = x;																p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																p[2].vt = VT_YPOS_PIXELS;

				// pack the pFilterButtonRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4C7D8A62-0F3A-41a7-BD58-FF25497D8451}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{00B2ADA1-29FD-45ce-A4C3-A2B85928FADD}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo;
				p[1].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pFilterButtonRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pFilterButtonRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pFilterButtonRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pFilterButtonRectangle->Top;

				p[0].pboolVal = pRaiseFilterChanged;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FILTERBUTTONCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[1].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FilterChanged event</em>
	///
	/// \param[in] pColumn The column header whose filter was changed. If \c NULL, all columns' filters
	///            were changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FilterChanged, ExplorerListView::Raise_FilterChanged,
	///       Fire_FilterButtonClick, Fire_EditChange
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FilterChanged, ExplorerListView::Raise_FilterChanged,
	///       Fire_FilterButtonClick, Fire_EditChange
	/// \endif
	HRESULT Fire_FilterChanged(IListViewColumn* pColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pColumn;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FILTERCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FindVirtualItem event</em>
	///
	/// \param[in] pItemToStartWith The item at which the search should start.
	/// \param[in] searchMode A value specifying the meaning of the \c pSearchFor parameter. Any of the
	///            values defined by the \c SearchModeConstants enumeration is valid.
	/// \param[in] pSearchFor The criterion that the returned item should fulfill. This parameter's
	///            format depends on the \c searchMode parameter:
	///            - \c smItemData An integer value.
	///            - \c smText A string value.
	///            - \c smPartialText A string value.
	///            - \c smNearestPosition An array containing two integer values. The first one specifies
	///              the x-coordinate, the second one the y-coordinate (both in pixels and relative to the
	///              control's upper-left corner).
	/// \param[in] searchDirection A value specifying the direction to search. Any of the values
	///            defined by the \c SearchDirectionConstants enumeration is valid. The client should ignore
	///            this parameter if the \c searchMode parameter is not set to \c smNearestPosition.
	/// \param[in] wrapAtLastItem If set to \c VARIANT_TRUE, the search should be continued with the first
	///            item if the last item is reached. The client should ignore this parameter if \c searchMode
	///            is set to \c smNearestPosition.
	/// \param[out] ppFoundItem Receives an item matching the specified characteristics. May receive \c NULL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FindVirtualItem, ExplorerListView::Raise_FindVirtualItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FindVirtualItem, ExplorerListView::Raise_FindVirtualItem
	/// \endif
	HRESULT Fire_FindVirtualItem(IListViewItem* pItemToStartWith, SearchModeConstants searchMode, VARIANT* pSearchFor, SearchDirectionConstants searchDirection, VARIANT_BOOL wrapAtLastItem, IListViewItem** ppFoundItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pItemToStartWith;
				p[4].lVal = static_cast<LONG>(searchMode);											p[4].vt = VT_I4;
				p[3].pvarVal = pSearchFor;																			p[3].vt = VT_VARIANT | VT_BYREF;
				p[2].lVal = static_cast<LONG>(searchDirection);									p[2].vt = VT_I4;
				p[1] = wrapAtLastItem;
				p[0].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppFoundItem);		p[0].vt = VT_DISPATCH | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FINDVIRTUALITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FooterItemClick event</em>
	///
	/// \param[in] pFooterItem The clicked footer item.
	/// \param[in] button The mouse buttons that were pressed during the click. Any of the values defined by
	///            VB's \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in twips) of the click's position relative to the control's upper-left
	///            corner.
	/// \param[in] y The y-coordinate (in twips) of the click's position relative to the control'supper-left
	///             corner.
	/// \param[in] hitTestDetails Specifies the part of the control that was clicked. Any of the
	///            values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pRemoveFooterArea If \c VARIANT_FALSE, the caller should remove the footer area;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FooterItemClick, ExplorerListView::Raise_FooterItemClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FooterItemClick, ExplorerListView::Raise_FooterItemClick
	/// \endif
	HRESULT Fire_FooterItemClick(IListViewFooterItem* pFooterItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pRemoveFooterArea)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pFooterItem;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pRemoveFooterArea;								p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FOOTERITEMCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeColumnData event</em>
	///
	/// \param[in] pColumn The column whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FreeColumnData, ExplorerListView::Raise_FreeColumnData,
	///       Fire_RemovingColumn, Fire_RemovedColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FreeColumnData, ExplorerListView::Raise_FreeColumnData,
	///       Fire_RemovingColumn, Fire_RemovedColumn
	/// \endif
	HRESULT Fire_FreeColumnData(IListViewColumn* pColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pColumn;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FREECOLUMNDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeFooterItemData event</em>
	///
	/// \param[in] pListItem The item whose associated data shall be freed.
	/// \param[in] itemData The data associated with the footer item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FreeFooterItemData,
	///       ExplorerListView::Raise_FreeFooterItemData
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FreeFooterItemData,
	///       ExplorerListView::Raise_FreeFooterItemData
	/// \endif
	HRESULT Fire_FreeFooterItemData(IListViewFooterItem* pFooterItem, LONG itemData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pFooterItem;
				p[0] = itemData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FREEFOOTERITEMDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeItemData event</em>
	///
	/// \param[in] pListItem The item whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::FreeItemData, ExplorerListView::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::FreeItemData, ExplorerListView::Raise_FreeItemData,
	///       Fire_RemovingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_FreeItemData(IListViewItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_FREEITEMDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GetSubItemControl event</em>
	///
	/// \param[in] pListSubItem The sub-item that the representation control is requested for.
	/// \param[in] controlKind The kind of representation control being requested. Any of the values defined
	///            by the \c SubItemControlKindConstants enumeration are valid.
	/// \param[in] editMode Specifies how the label-edit mode is being entered. Any of the values defined by
	///            the \c SubItemEditModeConstants enumeration is valid.
	/// \param[in,out] pSubItemControl The representation control to use for the specified sub-item. Any of
	///                the values defined by the \c SubItemControlConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GetSubItemControl,
	///       ExplorerListView::Raise_GetSubItemControl, Fire_ConfigureSubItemControl, Fire_CustomDraw
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GetSubItemControl,
	///       ExplorerListView::Raise_GetSubItemControl, Fire_ConfigureSubItemControl, Fire_CustomDraw
	/// \endif
	HRESULT Fire_GetSubItemControl(IListViewSubItem* pListSubItem, SubItemControlKindConstants controlKind, SubItemEditModeConstants editMode, SubItemControlConstants* pSubItemControl)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListSubItem;
				p[2] = controlKind;
				p[1].lVal = static_cast<LONG>(editMode);									p[1].vt = VT_I4;
				p[0].plVal = reinterpret_cast<PLONG>(pSubItemControl);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GETSUBITEMCONTROL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GroupAsynchronousDrawFailed event</em>
	///
	/// \param[in] pGroup The group whose image failed to be drawn.
	/// \param[in,out] pNotificationDetails A \c NMLVASYNCDRAWN struct holding the notification details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GroupAsynchronousDrawFailed,
	///       ExplorerListView::Raise_GroupAsynchronousDrawFailed, ExLVwLibU::FAILEDIMAGEDETAILS,
	///       Fire_ItemAsynchronousDrawFailed,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms670576.aspx">NMLVASYNCDRAWN</a>
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GroupAsynchronousDrawFailed,
	///       ExplorerListView::Raise_GroupAsynchronousDrawFailed, ExLVwLibA::FAILEDIMAGEDETAILS,
	///       Fire_ItemAsynchronousDrawFailed,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms670576.aspx">NMLVASYNCDRAWN</a>
	/// \endif
	HRESULT Fire_GroupAsynchronousDrawFailed(IListViewGroup* pGroup, NMLVASYNCDRAWN* pNotificationDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pGroup;
				p[2].lVal = static_cast<LONG>(pNotificationDetails->hr);												p[2].vt = VT_I4;
				p[1].plVal = reinterpret_cast<PLONG>(&pNotificationDetails->dwRetFlags);				p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = reinterpret_cast<PLONG>(&pNotificationDetails->iRetImageIndex);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pNotificationDetails->pimldp parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidFAILEDIMAGEDETAILS = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{2A0B5154-15A3-4bc1-8B39-1A6774C74CE3}");
					CLSIDFromString(clsid, &clsidFAILEDIMAGEDETAILS);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidFAILEDIMAGEDETAILS), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{AC07F5F6-E070-46e6-912E-9117249AB52E}");
					CLSIDFromString(clsid, &clsidFAILEDIMAGEDETAILS);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidFAILEDIMAGEDETAILS), &pRecordInfo)));
				#endif
				VariantInit(&p[3]);
				p[3].vt = VT_RECORD | VT_BYREF;
				p[3].pRecInfo = pRecordInfo;
				p[3].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hImageList = HandleToLong(pNotificationDetails->pimldp->himl);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->IconIndex = pNotificationDetails->pimldp->i;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hDC = HandleToLong(pNotificationDetails->pimldp->hdcDst);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionX = pNotificationDetails->pimldp->x;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionY = pNotificationDetails->pimldp->y;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawLeft = pNotificationDetails->pimldp->xBitmap;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawTop = pNotificationDetails->pimldp->yBitmap;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawWidth = pNotificationDetails->pimldp->cx;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawHeight = pNotificationDetails->pimldp->cy;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->BackColor = pNotificationDetails->pimldp->rgbBk;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->ForeColor = pNotificationDetails->pimldp->rgbFg;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingStyle = static_cast<ImageDrawingStyleConstants>(pNotificationDetails->pimldp->fStyle & ~ILD_OVERLAYMASK);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->OverlayIndex = (pNotificationDetails->pimldp->fStyle & ILD_OVERLAYMASK) >> 8;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->RasterOperationCode = pNotificationDetails->pimldp->dwRop;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingEffects = static_cast<ImageDrawingEffectsConstants>(pNotificationDetails->pimldp->fState);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->AlphaChannel = static_cast<BYTE>(pNotificationDetails->pimldp->Frame);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->EffectColor = pNotificationDetails->pimldp->crEffect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GROUPASYNCHRONOUSDRAWFAILED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				pNotificationDetails->pimldp->himl = static_cast<HIMAGELIST>(LongToHandle(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hImageList));
				pNotificationDetails->pimldp->i = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->IconIndex;
				pNotificationDetails->pimldp->hdcDst = static_cast<HDC>(LongToHandle(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hDC));
				pNotificationDetails->pimldp->x = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionX;
				pNotificationDetails->pimldp->y = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionY;
				pNotificationDetails->pimldp->xBitmap = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawLeft;
				pNotificationDetails->pimldp->yBitmap = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawTop;
				pNotificationDetails->pimldp->cx = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawWidth;
				pNotificationDetails->pimldp->cy = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawHeight;
				pNotificationDetails->pimldp->rgbBk = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->BackColor;
				pNotificationDetails->pimldp->rgbFg = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->ForeColor;
				pNotificationDetails->pimldp->fStyle = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingStyle | INDEXTOOVERLAYMASK(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->OverlayIndex);
				pNotificationDetails->pimldp->dwRop = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->RasterOperationCode;
				pNotificationDetails->pimldp->fState = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingEffects;
				pNotificationDetails->pimldp->Frame = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->AlphaChannel;
				pNotificationDetails->pimldp->crEffect = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->EffectColor;

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[3].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GroupCustomDraw event</em>
	///
	/// \param[in] pGroup The group that the notification refers to.
	/// \param[in,out] pTextColor An \c OLE_COLOR value specifying the color to draw the group's text in.
	///                The client may change this value.
	/// \param[in,out] pHeaderAlignment The alignment to draw the group's header text with. Any of the values
	///                defined by the \c AlignmentConstants enumeration is valid. The client may change this
	///                value.
	/// \param[in,out] pFooterAlignment The alignment to draw the group's footer text with. Any of the values
	///                defined by the \c AlignmentConstants enumeration is valid. The client may change this
	///                value.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Any of the values
	///            defined by the \c CustomDrawStageConstants enumeration is valid.
	/// \param[in] groupState The group's current state (focused, selected etc.). Most of the values
	///            defined by the \c CustomDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	/// \param[in] pTextRectangle A \c RECTANGLE structure specifying the bounding rectangle of the area,
	///            in which the group's text is to be drawn.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GroupCustomDraw,
	///       ExplorerListView::Raise_GroupCustomDraw, Fire_CustomDraw, Fire_HeaderCustomDraw
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GroupCustomDraw,
	///       ExplorerListView::Raise_GroupCustomDraw, Fire_CustomDraw, Fire_HeaderCustomDraw
	/// \endif
	HRESULT Fire_GroupCustomDraw(IListViewGroup* pGroup, OLE_COLOR* pTextColor, AlignmentConstants* pHeaderAlignment, AlignmentConstants* pFooterAlignment, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants groupState, LONG hDC, RECTANGLE* pDrawingRectangle, RECTANGLE* pTextRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pGroup;
				p[8].plVal = reinterpret_cast<PLONG>(pTextColor);						p[8].vt = VT_I4 | VT_BYREF;
				p[7].plVal = reinterpret_cast<PLONG>(pHeaderAlignment);			p[7].vt = VT_I4 | VT_BYREF;
				p[6].plVal = reinterpret_cast<PLONG>(pFooterAlignment);			p[6].vt = VT_I4 | VT_BYREF;
				p[5].lVal = static_cast<LONG>(drawStage);										p[5].vt = VT_I4;
				p[4].lVal = static_cast<LONG>(groupState);									p[4].vt = VT_I4;
				p[3] = hDC;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo1 = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4C7D8A62-0F3A-41a7-BD58-FF25497D8451}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
				#else
					LPOLESTR clsid = OLESTR("{00B2ADA1-29FD-45ce-A4C3-A2B85928FADD}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
				#endif
				VariantInit(&p[2]);
				p[2].vt = VT_RECORD | VT_BYREF;
				p[2].pRecInfo = pRecordInfo1;
				p[2].pvRecord = pRecordInfo1->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Top = pDrawingRectangle->Top;

				// pack the pTextRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo2 = NULL;
				#ifdef UNICODE
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#else
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo2;
				p[1].pvRecord = pRecordInfo2->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pTextRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pTextRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pTextRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pTextRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GROUPCUSTOMDRAW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo1) {
					pRecordInfo1->RecordDestroy(p[2].pvRecord);
				}
				if(pRecordInfo2) {
					pRecordInfo2->RecordDestroy(p[1].pvRecord);
				}
			}
		}

		/* Although pTextColor is of type OLE_COLOR, it may contain RGB colors only. So convert OLE_COLOR
		   colors into RGB colors. */
		if(*pTextColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pTextColor, NULL, &color);
			*pTextColor = static_cast<OLE_COLOR>(color);
		}

		return hr;
	}

	/// \brief <em>Fires the \c GroupGotFocus event</em>
	///
	/// \param[in] pGroup The group that has received the keyboard focus.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GroupGotFocus, ExplorerListView::Raise_GroupGotFocus,
	///       Fire_GroupSelectionChanged, Fire_GroupLostFocus
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GroupGotFocus, ExplorerListView::Raise_GroupGotFocus,
	///       Fire_GroupSelectionChanged, Fire_GroupLostFocus
	/// \endif
	HRESULT Fire_GroupGotFocus(IListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GROUPGOTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GroupLostFocus event</em>
	///
	/// \param[in] pGroup The group that has received the keyboard focus.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GroupLostFocus, ExplorerListView::Raise_GroupLostFocus,
	///       Fire_GroupGotFocus
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GroupLostFocus, ExplorerListView::Raise_GroupLostFocus,
	///       Fire_GroupGotFocus
	/// \endif
	HRESULT Fire_GroupLostFocus(IListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GROUPLOSTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GroupSelectionChanged event</em>
	///
	/// \param[in] pGroup The group that has been selected/unselected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GroupSelectionChanged,
	///       ExplorerListView::Raise_GroupSelectionChanged, Fire_GroupGotFocus, Fire_ItemSelectionChanged
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GroupSelectionChanged,
	///       ExplorerListView::Raise_GroupSelectionChanged, Fire_GroupGotFocus, Fire_ItemSelectionChanged
	/// \endif
	HRESULT Fire_GroupSelectionChanged(IListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GROUPSELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GroupTaskLinkClick event</em>
	///
	/// \param[in] pGroup The group whose task link was clicked.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined by the
	///            \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::GroupTaskLinkClick,
	///       ExplorerListView::Raise_GroupTaskLinkClick, Fire_Click
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::GroupTaskLinkClick,
	///       ExplorerListView::Raise_GroupTaskLinkClick, Fire_Click
	/// \endif
	HRESULT Fire_GroupTaskLinkClick(IListViewGroup* pGroup, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pGroup;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_GROUPTASKLINKCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderAbortedDrag event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderAbortedDrag,
	///       ExplorerListView::Raise_HeaderAbortedDrag, Fire_HeaderDrop, Fire_AbortedDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderAbortedDrag,
	///       ExplorerListView::Raise_HeaderAbortedDrag, Fire_HeaderDrop, Fire_AbortedDrag
	/// \endif
	HRESULT Fire_HeaderAbortedDrag(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERABORTEDDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderChevronClick event</em>
	///
	/// \param[in] pFirstOverflownColumn The first overflown column.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should not show the default overflow
	///                menu; otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderChevronClick,
	///       ExplorerListView::Raise_HeaderChevronClick, Fire_HeaderClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderChevronClick,
	///       ExplorerListView::Raise_HeaderChevronClick, Fire_HeaderClick
	/// \endif
	HRESULT Fire_HeaderChevronClick(IListViewColumn* pFirstOverflownColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pFirstOverflownColumn;
				p[4] = button;											p[4].vt = VT_I2;
				p[3] = shift;												p[3].vt = VT_I2;
				p[2] = x;														p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;														p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pShowDefaultMenu;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERCHEVRONCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderClick event</em>
	///
	/// \param[in] pColumn The clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was clicked. Any of the values defined
	///            by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderClick, ExplorerListView::Raise_HeaderClick,
	///       Fire_HeaderDblClick, Fire_HeaderMClick, Fire_HeaderRClick, Fire_HeaderXClick, Fire_ColumnClick,
	///       Fire_HeaderChevronClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderClick, ExplorerListView::Raise_HeaderClick,
	///       Fire_HeaderDblClick, Fire_HeaderMClick, Fire_HeaderRClick, Fire_HeaderXClick, Fire_ColumnClick,
	///       Fire_HeaderChevronClick
	/// \endif
	HRESULT Fire_HeaderClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderContextMenu event</em>
	///
	/// \param[in] pColumn The column header that the context menu refers to. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	/// \param[in,out] pShowDefaultMenu If \c VARIANT_FALSE, the caller should prevent the \c ShellListView
	///                control from showing the default context menu; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderContextMenu,
	///       ExplorerListView::Raise_HeaderContextMenu, Fire_HeaderRClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderContextMenu,
	///       ExplorerListView::Raise_HeaderContextMenu, Fire_HeaderRClick
	/// \endif
	HRESULT Fire_HeaderContextMenu(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails, VARIANT_BOOL* pShowDefaultMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pColumn;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pShowDefaultMenu;									p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERCONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderCustomDraw event</em>
	///
	/// \param[in] pColumn The column header that the notification refers to. May be \c NULL.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Most of the values
	///            defined by the \c CustomDrawStageConstants enumeration is valid.
	/// \param[in] columnState The column header's current state (focused, selected etc.). Some of the
	///            values defined by the \c CustomDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderCustomDraw,
	///       ExplorerListView::Raise_HeaderCustomDraw, Fire_CustomDraw, Fire_GroupCustomDraw,
	///       Fire_HeaderOwnerDrawItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderCustomDraw,
	///       ExplorerListView::Raise_HeaderCustomDraw, Fire_CustomDraw, Fire_GroupCustomDraw,
	///       Fire_HeaderOwnerDrawItem
	/// \endif
	HRESULT Fire_HeaderCustomDraw(IListViewColumn* pColumn, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants columnState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4].lVal = static_cast<LONG>(drawStage);										p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(columnState);									p[3].vt = VT_I4;
				p[2] = hDC;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4C7D8A62-0F3A-41a7-BD58-FF25497D8451}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{00B2ADA1-29FD-45ce-A4C3-A2B85928FADD}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo;
				p[1].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERCUSTOMDRAW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[1].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderDblClick event</em>
	///
	/// \param[in] pColumn The double-clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was double-clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderDblClick, ExplorerListView::Raise_HeaderDblClick,
	///       Fire_HeaderClick, Fire_HeaderMDblClick, Fire_HeaderRDblClick, Fire_HeaderXDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderDblClick, ExplorerListView::Raise_HeaderDblClick,
	///       Fire_HeaderClick, Fire_HeaderMDblClick, Fire_HeaderRDblClick, Fire_HeaderXDblClick
	/// \endif
	HRESULT Fire_HeaderDblClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderDividerDblClick event</em>
	///
	/// \param[in] pColumn The column header that the double-clicked column divider belongs to.
	/// \param[in,out] pAutoSizeColumn If \c VARIANT_TRUE, the caller should auto-size the specified column;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderDividerDblClick,
	///       ExplorerListView::Raise_HeaderDividerDblClick, Fire_HeaderDblClick, Fire_ResizingColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderDividerDblClick,
	///       ExplorerListView::Raise_HeaderDividerDblClick, Fire_HeaderDblClick, Fire_ResizingColumn
	/// \endif
	HRESULT Fire_HeaderDividerDblClick(IListViewColumn* pColumn, VARIANT_BOOL* pAutoSizeColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pColumn;
				p[0].pboolVal = pAutoSizeColumn;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERDIVIDERDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderDragMouseMove event</em>
	///
	/// \param[in,out] ppDropTarget The column header that is the current target of the drag'n'drop
	///                operation. The client may set this parameter to another column header.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] xListView The x-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] yListView The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderDragMouseMove,
	///       ExplorerListView::Raise_HeaderDragMouseMove, Fire_HeaderMouseMove, Fire_HeaderOLEDragMouseMove,
	///       Fire_DragMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderDragMouseMove,
	///       ExplorerListView::Raise_HeaderDragMouseMove, Fire_HeaderMouseMove, Fire_HeaderOLEDragMouseMove,
	///       Fire_DragMouseMove
	/// \endif
	HRESULT Fire_HeaderDragMouseMove(IListViewColumn** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, OLE_XPOS_PIXELS xListView, OLE_YPOS_PIXELS yListView, HeaderHitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[9].vt = VT_DISPATCH | VT_BYREF;
				p[8] = button;																									p[8].vt = VT_I2;
				p[7] = shift;																										p[7].vt = VT_I2;
				p[6] = x;																												p[6].vt = VT_XPOS_PIXELS;
				p[5] = y;																												p[5].vt = VT_YPOS_PIXELS;
				p[4] = xListView;																								p[4].vt = VT_XPOS_PIXELS;
				p[3] = yListView;																								p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderDrop event</em>
	///
	/// \param[in] pDropTarget The column header that is the nearest one from the mouse cursor's position.
	///            If the mouse cursor's position lies outside the header control's client area (the
	///            y-coordinate may also lie some pixels above or below the client area), this parameter
	///            should be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that the mouse cursor's position lies in.
	///            Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderDrop, ExplorerListView::Raise_HeaderDrop,
	///       Fire_HeaderAbortedDrag, Fire_Drop
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderDrop, ExplorerListView::Raise_HeaderDrop,
	///       Fire_HeaderAbortedDrag, Fire_Drop
	/// \endif
	HRESULT Fire_HeaderDrop(IListViewColumn* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderGotFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderGotFocus, ExplorerListView::Raise_HeaderGotFocus,
	///       Fire_HeaderLostFocus
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderGotFocus, ExplorerListView::Raise_HeaderGotFocus,
	///       Fire_HeaderLostFocus
	/// \endif
	HRESULT Fire_HeaderGotFocus(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERGOTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderItemGetDisplayInfo event</em>
	///
	/// \param[in] pColumn The column header that the value is required for.
	/// \param[in] requestedInfo Specifies which properties' values are required. Some combinations of
	///            the values defined by the \c RequestedInfoConstants enumeration are valid.
	/// \param[out] pIconIndex The zero-based index of the requested icon. If the \c requestedInfo
	///             parameter doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[in] maxColumnCaptionLength The maximum number of characters the column header's text may
	///            consist of. If the \c requestedInfo parameter doesn't include \c riItemText, the client
	///            should ignore this value.
	/// \param[out] pColumnCaption The column header's caption. If the \c requestedInfo parameter doesn't
	///             include \c riItemText, the caller should ignore this value.
	/// \param[in,out] pDontAskAgain If \c VARIANT_TRUE, the caller should always use the same settings
	///                and never fire this event again for these properties of this column header;
	///                otherwise it shouldn't make the values persistent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderItemGetDisplayInfo,
	///       ExplorerListView::Raise_HeaderItemGetDisplayInfo
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderItemGetDisplayInfo,
	///       ExplorerListView::Raise_HeaderItemGetDisplayInfo
	/// \endif
	HRESULT Fire_HeaderItemGetDisplayInfo(IListViewColumn* pColumn, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG maxColumnCaptionLength, BSTR* pColumnCaption, VARIANT_BOOL* pDontAskAgain)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = static_cast<LONG>(requestedInfo);		p[4].vt = VT_I4;
				p[3].plVal = pIconIndex;										p[3].vt = VT_I4 | VT_BYREF;
				p[2] = maxColumnCaptionLength;
				p[1].pbstrVal = pColumnCaption;							p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pDontAskAgain;							p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERITEMGETDISPLAYINFO, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderKeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderKeyDown, ExplorerListView::Raise_HeaderKeyDown,
	///       Fire_HeaderKeyUp, Fire_HeaderKeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderKeyDown, ExplorerListView::Raise_HeaderKeyDown,
	///       Fire_HeaderKeyUp, Fire_HeaderKeyPress
	/// \endif
	HRESULT Fire_HeaderKeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERKEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderKeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderKeyPress, ExplorerListView::Raise_HeaderKeyPress,
	///       Fire_HeaderKeyDown, Fire_HeaderKeyUp
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderKeyPress, ExplorerListView::Raise_HeaderKeyPress,
	///       Fire_HeaderKeyDown, Fire_HeaderKeyUp
	/// \endif
	HRESULT Fire_HeaderKeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERKEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderKeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderKeyUp, ExplorerListView::Raise_HeaderKeyUp,
	///       Fire_HeaderKeyDown, Fire_HeaderKeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderKeyUp, ExplorerListView::Raise_HeaderKeyUp,
	///       Fire_HeaderKeyDown, Fire_HeaderKeyPress
	/// \endif
	HRESULT Fire_HeaderKeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERKEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderLostFocus event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderLostFocus, ExplorerListView::Raise_HeaderLostFocus,
	///       Fire_HeaderGotFocus
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderLostFocus, ExplorerListView::Raise_HeaderLostFocus,
	///       Fire_HeaderGotFocus
	/// \endif
	HRESULT Fire_HeaderLostFocus(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERLOSTFOCUS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMClick event</em>
	///
	/// \param[in] pColumn The clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMClick, ExplorerListView::Raise_HeaderMClick,
	///       Fire_HeaderMDblClick, Fire_HeaderClick, Fire_HeaderRClick, Fire_HeaderXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMClick, ExplorerListView::Raise_HeaderMClick,
	///       Fire_HeaderMDblClick, Fire_HeaderClick, Fire_HeaderRClick, Fire_HeaderXClick
	/// \endif
	HRESULT Fire_HeaderMClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMDblClick event</em>
	///
	/// \param[in] pColumn The double-clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was double-clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMDblClick, ExplorerListView::Raise_HeaderMDblClick,
	///       Fire_HeaderMClick, Fire_HeaderDblClick, Fire_HeaderRDblClick, Fire_HeaderXDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMDblClick, ExplorerListView::Raise_HeaderMDblClick,
	///       Fire_HeaderMClick, Fire_HeaderDblClick, Fire_HeaderRDblClick, Fire_HeaderXDblClick
	/// \endif
	HRESULT Fire_HeaderMDblClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseDown event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseDown, ExplorerListView::Raise_HeaderMouseDown,
	///       Fire_HeaderMouseUp, Fire_HeaderClick, Fire_HeaderMClick, Fire_HeaderRClick, Fire_HeaderXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseDown, ExplorerListView::Raise_HeaderMouseDown,
	///       Fire_HeaderMouseUp, Fire_HeaderClick, Fire_HeaderMClick, Fire_HeaderRClick, Fire_HeaderXClick
	/// \endif
	HRESULT Fire_HeaderMouseDown(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseEnter event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseEnter, ExplorerListView::Raise_HeaderMouseEnter,
	///       Fire_HeaderMouseLeave, Fire_ColumnMouseEnter, Fire_HeaderMouseHover, Fire_HeaderMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseEnter, ExplorerListView::Raise_HeaderMouseEnter,
	///       Fire_HeaderMouseLeave, Fire_ColumnMouseEnter, Fire_HeaderMouseHover, Fire_HeaderMouseMove
	/// \endif
	HRESULT Fire_HeaderMouseEnter(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseHover event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseHover, ExplorerListView::Raise_HeaderMouseHover,
	///       Fire_HeaderMouseEnter, Fire_HeaderMouseLeave, Fire_HeaderMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseHover, ExplorerListView::Raise_HeaderMouseHover,
	///       Fire_HeaderMouseEnter, Fire_HeaderMouseLeave, Fire_HeaderMouseMove
	/// \endif
	HRESULT Fire_HeaderMouseHover(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseLeave event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseLeave, ExplorerListView::Raise_HeaderMouseLeave,
	///       Fire_HeaderMouseEnter, Fire_ColumnMouseLeave, Fire_HeaderMouseHover, Fire_HeaderMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseLeave, ExplorerListView::Raise_HeaderMouseLeave,
	///       Fire_HeaderMouseEnter, Fire_ColumnMouseLeave, Fire_HeaderMouseHover, Fire_HeaderMouseMove
	/// \endif
	HRESULT Fire_HeaderMouseLeave(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseMove event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseMove, ExplorerListView::Raise_HeaderMouseMove,
	///       Fire_HeaderMouseEnter, Fire_HeaderMouseLeave, Fire_HeaderMouseDown, Fire_HeaderMouseUp,
	///       Fire_HeaderMouseWheel
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseMove, ExplorerListView::Raise_HeaderMouseMove,
	///       Fire_HeaderMouseEnter, Fire_HeaderMouseLeave, Fire_HeaderMouseDown, Fire_HeaderMouseUp,
	///       Fire_HeaderMouseWheel
	/// \endif
	HRESULT Fire_HeaderMouseMove(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseUp event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseUp, ExplorerListView::Raise_HeaderMouseUp,
	///       Fire_HeaderMouseDown, Fire_HeaderClick, Fire_HeaderMClick, Fire_HeaderRClick, Fire_HeaderXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseUp, ExplorerListView::Raise_HeaderMouseUp,
	///       Fire_HeaderMouseDown, Fire_HeaderClick, Fire_HeaderMClick, Fire_HeaderRClick, Fire_HeaderXClick
	/// \endif
	HRESULT Fire_HeaderMouseUp(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderMouseWheel event</em>
	///
	/// \param[in] pColumn The column header that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderMouseWheel,
	///       ExplorerListView::Raise_HeaderMouseWheel, Fire_HeaderMouseMove, Fire_MouseWheel,
	///       Fire_EditMouseWheel
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderMouseWheel,
	///       ExplorerListView::Raise_HeaderMouseWheel, Fire_HeaderMouseMove, Fire_MouseWheel,
	///       Fire_EditMouseWheel
	/// \endif
	HRESULT Fire_HeaderMouseWheel(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pColumn;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);		p[2].vt = VT_I4;
				p[1].lVal = static_cast<LONG>(scrollAxis);				p[1].vt = VT_I4;
				p[0] = wheelDelta;																p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERMOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLECompleteDrag,
	///       ExplorerListView::Raise_HeaderOLECompleteDrag, Fire_HeaderOLEStartDrag, Fire_OLECompleteDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLECompleteDrag,
	///       ExplorerListView::Raise_HeaderOLECompleteDrag, Fire_HeaderOLEStartDrag, Fire_OLECompleteDrag
	/// \endif
	HRESULT Fire_HeaderOLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pData;
				p[0].lVal = static_cast<LONG>(performedEffect);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in] pDropTarget The column header object that is the nearest one from the mouse cursor's
	///            position.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] xListView The x-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] yListView The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEDragDrop,
	///       ExplorerListView::Raise_HeaderOLEDragDrop, Fire_HeaderOLEDragEnter,
	///       Fire_HeaderOLEDragMouseMove, Fire_HeaderOLEDragLeave, Fire_HeaderMouseUp, Fire_OLEDragDrop
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEDragDrop,
	///       ExplorerListView::Raise_HeaderOLEDragDrop, Fire_HeaderOLEDragEnter,
	///       Fire_HeaderOLEDragMouseMove, Fire_HeaderOLEDragLeave, Fire_HeaderMouseUp, Fire_OLEDragDrop
	/// \endif
	HRESULT Fire_HeaderOLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListViewColumn* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, OLE_XPOS_PIXELS xListView, OLE_YPOS_PIXELS yListView, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pData;
				p[8].plVal = reinterpret_cast<PLONG>(pEffect);		p[8].vt = VT_I4 | VT_BYREF;
				p[7] = pDropTarget;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2] = xListView;																	p[2].vt = VT_XPOS_PIXELS;
				p[1] = yListView;																	p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The column header that is the current target of the drag'n'drop
	///                operation. The client may set this parameter to another column header.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] xListView The x-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] yListView The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEDragEnter,
	///       ExplorerListView::Raise_HeaderOLEDragEnter, Fire_HeaderOLEDragMouseMove,
	///       Fire_HeaderOLEDragLeave, Fire_HeaderOLEDragDrop, Fire_HeaderMouseEnter, Fire_OLEDragEnter
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEDragEnter,
	///       ExplorerListView::Raise_HeaderOLEDragEnter, Fire_HeaderOLEDragMouseMove,
	///       Fire_HeaderOLEDragLeave, Fire_HeaderOLEDragDrop, Fire_HeaderMouseEnter, Fire_OLEDragEnter
	/// \endif
	HRESULT Fire_HeaderOLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListViewColumn** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, OLE_XPOS_PIXELS xListView, OLE_YPOS_PIXELS yListView, HeaderHitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[12];
				p[11] = pData;
				p[10].plVal = reinterpret_cast<PLONG>(pEffect);									p[10].vt = VT_I4 | VT_BYREF;
				p[9].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[9].vt = VT_DISPATCH | VT_BYREF;
				p[8] = button;																									p[8].vt = VT_I2;
				p[7] = shift;																										p[7].vt = VT_I2;
				p[6] = x;																												p[6].vt = VT_XPOS_PIXELS;
				p[5] = y;																												p[5].vt = VT_YPOS_PIXELS;
				p[4] = xListView;																								p[4].vt = VT_XPOS_PIXELS;
				p[3] = yListView;																								p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 12, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The handle of the potential drag'n'drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEDragEnterPotentialTarget,
	///       ExplorerListView::Raise_HeaderOLEDragEnterPotentialTarget,
	///       Fire_HeaderOLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEDragEnterPotentialTarget,
	///       ExplorerListView::Raise_HeaderOLEDragEnterPotentialTarget,
	///       Fire_HeaderOLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \endif
	HRESULT Fire_HeaderOLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndPotentialTarget;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The column header that is the current target of the drag'n'drop operation.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] xListView The x-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] yListView The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEDragLeave,
	///       ExplorerListView::Raise_HeaderOLEDragLeave, Fire_HeaderOLEDragEnter,
	///       Fire_HeaderOLEDragMouseMove, Fire_HeaderOLEDragDrop, Fire_HeaderMouseLeave, Fire_OLEDragLeave
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEDragLeave,
	///       ExplorerListView::Raise_HeaderOLEDragLeave, Fire_HeaderOLEDragEnter,
	///       Fire_HeaderOLEDragMouseMove, Fire_HeaderOLEDragDrop, Fire_HeaderMouseLeave, Fire_OLEDragLeave
	/// \endif
	HRESULT Fire_HeaderOLEDragLeave(IOLEDataObject* pData, IListViewColumn* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, OLE_XPOS_PIXELS xListView, OLE_YPOS_PIXELS yListView, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7] = pDropTarget;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2] = xListView;																	p[2].vt = VT_XPOS_PIXELS;
				p[1] = yListView;																	p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEDragLeavePotentialTarget,
	///       ExplorerListView::Raise_HeaderOLEDragLeavePotentialTarget,
	///       Fire_HeaderOLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEDragLeavePotentialTarget,
	///       ExplorerListView::Raise_HeaderOLEDragLeavePotentialTarget,
	///       Fire_HeaderOLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \endif
	HRESULT Fire_HeaderOLEDragLeavePotentialTarget(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The column header that is the current target of the drag'n'drop
	///                operation. The client may set this parameter to another column header.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the header
	///            control's upper-left corner.
	/// \param[in] xListView The x-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] yListView The y-coordinate (in pixels) of the mouse cursor's position relative to the
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The exact part of the header control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HeaderHitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEDragMouseMove,
	///       ExplorerListView::Raise_HeaderOLEDragMouseMove, Fire_HeaderOLEDragEnter,
	///       Fire_HeaderOLEDragLeave, Fire_HeaderOLEDragDrop, Fire_HeaderMouseMove, Fire_OLEDragMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEDragMouseMove,
	///       ExplorerListView::Raise_HeaderOLEDragMouseMove, Fire_HeaderOLEDragEnter,
	///       Fire_HeaderOLEDragLeave, Fire_HeaderOLEDragDrop, Fire_HeaderMouseMove, Fire_OLEDragMouseMove
	/// \endif
	HRESULT Fire_HeaderOLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListViewColumn** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, OLE_XPOS_PIXELS xListView, OLE_YPOS_PIXELS yListView, HeaderHitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[12];
				p[11] = pData;
				p[10].plVal = reinterpret_cast<PLONG>(pEffect);									p[10].vt = VT_I4 | VT_BYREF;
				p[9].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[9].vt = VT_DISPATCH | VT_BYREF;
				p[8] = button;																									p[8].vt = VT_I2;
				p[7] = shift;																										p[7].vt = VT_I2;
				p[6] = x;																												p[6].vt = VT_XPOS_PIXELS;
				p[5] = y;																												p[5].vt = VT_YPOS_PIXELS;
				p[4] = xListView;																								p[4].vt = VT_XPOS_PIXELS;
				p[3] = yListView;																								p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 12, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c OLEDropEffectConstants enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEGiveFeedback,
	///       ExplorerListView::Raise_HeaderOLEGiveFeedback, Fire_HeaderOLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEGiveFeedback,
	///       ExplorerListView::Raise_HeaderOLEGiveFeedback, Fire_HeaderOLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \endif
	HRESULT Fire_HeaderOLEGiveFeedback(OLEDropEffectConstants effect, VARIANT_BOOL* pUseDefaultCursors)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = static_cast<LONG>(effect);			p[1].vt = VT_I4;
				p[0].pboolVal = pUseDefaultCursors;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c VARIANT_TRUE, the user has pressed the \c ESC key since the last
	///            time this event was fired; otherwise not.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue, cancel or complete the
	///                drag'n'drop operation. Any of the values defined by the
	///                \c OLEActionToContinueWithConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEQueryContinueDrag,
	///       ExplorerListView::Raise_HeaderOLEQueryContinueDrag, Fire_HeaderOLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEQueryContinueDrag,
	///       ExplorerListView::Raise_HeaderOLEQueryContinueDrag, Fire_HeaderOLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \endif
	HRESULT Fire_HeaderOLEQueryContinueDrag(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, OLEActionToContinueWithConstants* pActionToContinueWith)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pressedEscape;
				p[2] = button;																									p[2].vt = VT_I2;
				p[1] = shift;																										p[1].vt = VT_I2;
				p[0].plVal = reinterpret_cast<PLONG>(pActionToContinueWith);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEReceivedNewData,
	///       ExplorerListView::Raise_HeaderOLEReceivedNewData, Fire_OLEReceivedNewData,
	///       Fire_HeaderOLESetData
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEReceivedNewData,
	///       ExplorerListView::Raise_HeaderOLEReceivedNewData, Fire_OLEReceivedNewData,
	///       Fire_HeaderOLESetData
	/// \endif
	HRESULT Fire_HeaderOLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLESetData,
	///       ExplorerListView::Raise_HeaderOLESetData, Fire_HeaderOLEStartDrag, Fire_OLESetData
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLESetData,
	///       ExplorerListView::Raise_HeaderOLESetData, Fire_HeaderOLEStartDrag, Fire_OLESetData
	/// \endif
	HRESULT Fire_HeaderOLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOLEStartDrag,
	///       ExplorerListView::Raise_HeaderOLEStartDrag, Fire_HeaderOLESetData, Fire_HeaderOLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOLEStartDrag,
	///       ExplorerListView::Raise_HeaderOLEStartDrag, Fire_HeaderOLESetData, Fire_HeaderOLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_HeaderOLEStartDrag(IOLEDataObject* pData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderOwnerDrawItem event</em>
	///
	/// \param[in] pColumn The column header to draw.
	/// \param[in] columnState The column header's current state (focused, selected etc.). Some of the values
	///            defined by the \c OwnerDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderOwnerDrawItem,
	///       ExplorerListView::Raise_HeaderOwnerDrawItem, Fire_HeaderCustomDraw, Fire_OwnerDrawItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderOwnerDrawItem,
	///       ExplorerListView::Raise_HeaderOwnerDrawItem, Fire_HeaderCustomDraw, Fire_OwnerDrawItem
	/// \endif
	HRESULT Fire_HeaderOwnerDrawItem(IListViewColumn* pColumn, OwnerDrawItemStateConstants columnState, LONG hDC, RECTANGLE* pDrawingRectangle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pColumn;
				p[2].lVal = static_cast<LONG>(columnState);		p[2].vt = VT_I4;
				p[1] = hDC;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4C7D8A62-0F3A-41a7-BD58-FF25497D8451}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{00B2ADA1-29FD-45ce-A4C3-A2B85928FADD}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[0]);
				p[0].vt = VT_RECORD | VT_BYREF;
				p[0].pRecInfo = pRecordInfo;
				p[0].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADEROWNERDRAWITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[0].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderRClick event</em>
	///
	/// \param[in] pColumn The clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderRClick, ExplorerListView::Raise_HeaderRClick,
	///       Fire_HeaderContextMenu, Fire_HeaderRDblClick, Fire_HeaderClick, Fire_HeaderMClick,
	///       Fire_HeaderXClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderRClick, ExplorerListView::Raise_HeaderRClick,
	///       Fire_HeaderContextMenu, Fire_HeaderRDblClick, Fire_HeaderClick, Fire_HeaderMClick,
	///       Fire_HeaderXClick
	/// \endif
	HRESULT Fire_HeaderRClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERRCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderRDblClick event</em>
	///
	/// \param[in] pColumn The double-clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was double-clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderRDblClick, ExplorerListView::Raise_HeaderRDblClick,
	///       Fire_HeaderRClick, Fire_HeaderDblClick, Fire_HeaderMDblClick, Fire_HeaderXDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderRDblClick, ExplorerListView::Raise_HeaderRDblClick,
	///       Fire_HeaderRClick, Fire_HeaderDblClick, Fire_HeaderMDblClick, Fire_HeaderXDblClick
	/// \endif
	HRESULT Fire_HeaderRDblClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERRDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderXClick event</em>
	///
	/// \param[in] pColumn The clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the header control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderXClick, ExplorerListView::Raise_HeaderXClick,
	///       Fire_HeaderXDblClick, Fire_HeaderClick, Fire_HeaderMClick, Fire_HeaderRClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderXClick, ExplorerListView::Raise_HeaderXClick,
	///       Fire_HeaderXDblClick, Fire_HeaderClick, Fire_HeaderMClick, Fire_HeaderRClick
	/// \endif
	HRESULT Fire_HeaderXClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERXCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeaderXDblClick event</em>
	///
	/// \param[in] pColumn The double-clicked column header. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the header
	///            control's upper-left corner.
	/// \param[in] hitTestDetails The part of the header control that was double-clicked. Any of the values
	///            defined by the \c HeaderHitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HeaderXDblClick, ExplorerListView::Raise_HeaderXDblClick,
	///       Fire_HeaderXClick, Fire_HeaderDblClick, Fire_HeaderMDblClick, Fire_HeaderRDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HeaderXDblClick, ExplorerListView::Raise_HeaderXDblClick,
	///       Fire_HeaderXClick, Fire_HeaderDblClick, Fire_HeaderMDblClick, Fire_HeaderRDblClick
	/// \endif
	HRESULT Fire_HeaderXDblClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pColumn;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HEADERXDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HotItemChanged event</em>
	///
	/// \param[in] pPreviousHotItem The previous hot item. May be \ NULL.
	/// \param[in] pNewHotItem The new hot item. May be \ NULL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HotItemChanged, ExplorerListView::Raise_HotItemChanged,
	///       Fire_HotItemChanging
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HotItemChanged, ExplorerListView::Raise_HotItemChanged,
	///       Fire_HotItemChanging
	/// \endif
	HRESULT Fire_HotItemChanged(IListViewItem* pPreviousHotItem, IListViewItem* pNewHotItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pPreviousHotItem;
				p[0] = pNewHotItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HOTITEMCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HotItemChanging event</em>
	///
	/// \param[in] pPreviousHotItem The previous hot item. May be \ NULL.
	/// \param[in] pNewHotItem The new hot item. May be \ NULL.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort the hot item change;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::HotItemChanging, ExplorerListView::Raise_HotItemChanging,
	///       Fire_HotItemChanged
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::HotItemChanging, ExplorerListView::Raise_HotItemChanging,
	///       Fire_HotItemChanged
	/// \endif
	HRESULT Fire_HotItemChanging(IListViewItem* pPreviousHotItem, IListViewItem* pNewHotItem, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pPreviousHotItem;
				p[1] = pNewHotItem;
				p[0].pboolVal = pCancelChange;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_HOTITEMCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c IncrementalSearching event</em>
	///
	/// \param[in] currentSearchString The control's current incremental search-string.
	/// \param[in,out] pItemToSelect Receives the zero-based index of the item to select. If set to -1, the
	///                control searches an appropriate item based on the search string. If set to -2, the
	///                search is aborted with an error beep. If set to -3, incremental search string
	///                processing is stopped.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::IncrementalSearching,
	///       ExplorerListView::Raise_IncrementalSearching, Fire_IncrementalSearchStringChanging
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::IncrementalSearching,
	///       ExplorerListView::Raise_IncrementalSearching, Fire_IncrementalSearchStringChanging
	/// \endif
	HRESULT Fire_IncrementalSearching(BSTR currentSearchString, LONG* pItemToSelect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = currentSearchString;
				p[0].plVal = pItemToSelect;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INCREMENTALSEARCHING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c IncrementalSearchStringChanging event</em>
	///
	/// \param[in] currentSearchString The control's current incremental search-string.
	/// \param[in] keyCodeOfCharToBeAdded The key code of the character to be added to the search-string.
	///            Most of the values defined by VB's \c KeyCodeConstants enumeration are valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should discard the character and
	///                leave the search-string unchanged; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::IncrementalSearchStringChanging,
	///       ExplorerListView::Raise_IncrementalSearchStringChanging, Fire_KeyPress,
	///       Fire_IncrementalSearching
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::IncrementalSearchStringChanging,
	///       ExplorerListView::Raise_IncrementalSearchStringChanging, Fire_KeyPress,
	///       Fire_IncrementalSearching
	/// \endif
	HRESULT Fire_IncrementalSearchStringChanging(BSTR currentSearchString, SHORT keyCodeOfCharToBeAdded, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = currentSearchString;
				p[1] = keyCodeOfCharToBeAdded;
				p[0].pboolVal = pCancelChange;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INCREMENTALSEARCHSTRINGCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedColumn event</em>
	///
	/// \param[in] pColumn The inserted column.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InsertedColumn, ExplorerListView::Raise_InsertedColumn,
	///       Fire_InsertingColumn, Fire_RemovedColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InsertedColumn, ExplorerListView::Raise_InsertedColumn,
	///       Fire_InsertingColumn, Fire_RemovedColumn
	/// \endif
	HRESULT Fire_InsertedColumn(IListViewColumn* pColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pColumn;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INSERTEDCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedGroup event</em>
	///
	/// \param[in] pGroup The inserted group.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InsertedGroup, ExplorerListView::Raise_InsertedGroup,
	///       Fire_InsertingGroup, Fire_RemovedGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InsertedGroup, ExplorerListView::Raise_InsertedGroup,
	///       Fire_InsertingGroup, Fire_RemovedGroup
	/// \endif
	HRESULT Fire_InsertedGroup(IListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INSERTEDGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedItem event</em>
	///
	/// \param[in] pListItem The inserted item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InsertedItem, ExplorerListView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_RemovedItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InsertedItem, ExplorerListView::Raise_InsertedItem,
	///       Fire_InsertingItem, Fire_RemovedItem
	/// \endif
	HRESULT Fire_InsertedItem(IListViewItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INSERTEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingColumn event</em>
	///
	/// \param[in] pColumn The column being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InsertingColumn, ExplorerListView::Raise_InsertingColumn,
	///       Fire_InsertedColumn, Fire_RemovingColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InsertingColumn, ExplorerListView::Raise_InsertingColumn,
	///       Fire_InsertedColumn, Fire_RemovingColumn
	/// \endif
	HRESULT Fire_InsertingColumn(IVirtualListViewColumn* pColumn, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pColumn;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INSERTINGCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingGroup event</em>
	///
	/// \param[in] pGroup The group being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InsertingGroup, ExplorerListView::Raise_InsertingGroup,
	///       Fire_InsertedGroup, Fire_RemovingGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InsertingGroup, ExplorerListView::Raise_InsertingGroup,
	///       Fire_InsertedGroup, Fire_RemovingGroup
	/// \endif
	HRESULT Fire_InsertingGroup(IVirtualListViewGroup* pGroup, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pGroup;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INSERTINGGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingItem event</em>
	///
	/// \param[in] pListItem The item being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InsertingItem, ExplorerListView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InsertingItem, ExplorerListView::Raise_InsertingItem,
	///       Fire_InsertedItem, Fire_RemovingItem
	/// \endif
	HRESULT Fire_InsertingItem(IVirtualListViewItem* pListItem, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INSERTINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InvokeVerbFromSubItemControl event</em>
	///
	/// \param[in] pListItem The item on which the action has been invoked.
	/// \param[in] verb A string representing the invoked action. For the \c sicHyperlink sub-item control
	///            this is the value of the \c id attribute of the hyperlink tag that specifies the link.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::InvokeVerbFromSubItemControl,
	///       ExplorerListView::Raise_InvokeVerbFromSubItemControl, Fire_GetSubItemControl,
	///       Fire_ConfigureSubItemControl
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::InvokeVerbFromSubItemControl,
	///       ExplorerListView::Raise_InvokeVerbFromSubItemControl, Fire_GetSubItemControl,
	///       Fire_ConfigureSubItemControl
	/// \endif
	HRESULT Fire_InvokeVerbFromSubItemControl(IListViewItem* pListItem, BSTR verb)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0] = verb;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_INVOKEVERBFROMSUBITEMCONTROL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemActivate event</em>
	///
	/// \param[in] pListItem The item being activated.
	/// \param[in] pListSubItem The sub-item being activated. May be \c NULL.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner. Will be -1 if the item is activated using the keyboard instead of the
	///            mouse.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner. Will be -1 if the item is activated using the keyboard instead of the
	///            mouse.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemActivate, ExplorerListView::Raise_ItemActivate
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemActivate, ExplorerListView::Raise_ItemActivate
	/// \endif
	HRESULT Fire_ItemActivate(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT shift, LONG x, LONG y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pListItem;
				p[3] = pListSubItem;
				p[2] = shift;						p[2].vt = VT_I2;
				p[1] = x;
				p[0] = y;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMACTIVATE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemAsynchronousDrawFailed event</em>
	///
	/// \param[in] pListItem The item whose image failed to be drawn.
	/// \param[in] pListSubItem The sub-item whose image failed to be drawn. May be \c NULL.
	/// \param[in,out] pNotificationDetails A \c NMLVASYNCDRAWN struct holding the notification details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemAsynchronousDrawFailed,
	///       ExplorerListView::Raise_ItemAsynchronousDrawFailed, ExLVwLibU::FAILEDIMAGEDETAILS,
	///       Fire_GroupAsynchronousDrawFailed,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms670576.aspx">NMLVASYNCDRAWN</a>
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemAsynchronousDrawFailed,
	///       ExplorerListView::Raise_ItemAsynchronousDrawFailed, ExLVwLibA::FAILEDIMAGEDETAILS,
	///       Fire_GroupAsynchronousDrawFailed,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms670576.aspx">NMLVASYNCDRAWN</a>
	/// \endif
	HRESULT Fire_ItemAsynchronousDrawFailed(IListViewItem* pListItem, IListViewSubItem* pListSubItem, NMLVASYNCDRAWN* pNotificationDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = pListSubItem;
				p[2].lVal = static_cast<LONG>(pNotificationDetails->hr);												p[2].vt = VT_I4;
				p[1].plVal = reinterpret_cast<PLONG>(&pNotificationDetails->dwRetFlags);				p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = reinterpret_cast<PLONG>(&pNotificationDetails->iRetImageIndex);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pNotificationDetails->pimldp parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidFAILEDIMAGEDETAILS = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{2A0B5154-15A3-4bc1-8B39-1A6774C74CE3}");
					CLSIDFromString(clsid, &clsidFAILEDIMAGEDETAILS);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidFAILEDIMAGEDETAILS), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{AC07F5F6-E070-46e6-912E-9117249AB52E}");
					CLSIDFromString(clsid, &clsidFAILEDIMAGEDETAILS);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidFAILEDIMAGEDETAILS), &pRecordInfo)));
				#endif
				VariantInit(&p[3]);
				p[3].vt = VT_RECORD | VT_BYREF;
				p[3].pRecInfo = pRecordInfo;
				p[3].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hImageList = HandleToLong(pNotificationDetails->pimldp->himl);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->IconIndex = pNotificationDetails->pimldp->i;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hDC = HandleToLong(pNotificationDetails->pimldp->hdcDst);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionX = pNotificationDetails->pimldp->x;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionY = pNotificationDetails->pimldp->y;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawLeft = pNotificationDetails->pimldp->xBitmap;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawTop = pNotificationDetails->pimldp->yBitmap;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawWidth = pNotificationDetails->pimldp->cx;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawHeight = pNotificationDetails->pimldp->cy;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->BackColor = pNotificationDetails->pimldp->rgbBk;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->ForeColor = pNotificationDetails->pimldp->rgbFg;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingStyle = static_cast<ImageDrawingStyleConstants>(pNotificationDetails->pimldp->fStyle & ~ILD_OVERLAYMASK);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->OverlayIndex = (pNotificationDetails->pimldp->fStyle & ILD_OVERLAYMASK) >> 8;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->RasterOperationCode = pNotificationDetails->pimldp->dwRop;
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingEffects = static_cast<ImageDrawingEffectsConstants>(pNotificationDetails->pimldp->fState);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->AlphaChannel = static_cast<BYTE>(pNotificationDetails->pimldp->Frame);
				reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->EffectColor = pNotificationDetails->pimldp->crEffect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMASYNCHRONOUSDRAWFAILED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				pNotificationDetails->pimldp->himl = static_cast<HIMAGELIST>(LongToHandle(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hImageList));
				pNotificationDetails->pimldp->i = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->IconIndex;
				pNotificationDetails->pimldp->hdcDst = static_cast<HDC>(LongToHandle(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->hDC));
				pNotificationDetails->pimldp->x = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionX;
				pNotificationDetails->pimldp->y = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->TargetPositionY;
				pNotificationDetails->pimldp->xBitmap = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawLeft;
				pNotificationDetails->pimldp->yBitmap = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawTop;
				pNotificationDetails->pimldp->cx = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawWidth;
				pNotificationDetails->pimldp->cy = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->SectionToDrawHeight;
				pNotificationDetails->pimldp->rgbBk = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->BackColor;
				pNotificationDetails->pimldp->rgbFg = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->ForeColor;
				pNotificationDetails->pimldp->fStyle = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingStyle | INDEXTOOVERLAYMASK(reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->OverlayIndex);
				pNotificationDetails->pimldp->dwRop = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->RasterOperationCode;
				pNotificationDetails->pimldp->fState = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->DrawingEffects;
				pNotificationDetails->pimldp->Frame = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->AlphaChannel;
				pNotificationDetails->pimldp->crEffect = reinterpret_cast<FAILEDIMAGEDETAILS*>(p[3].pvRecord)->EffectColor;

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[3].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginDrag event</em>
	///
	/// \param[in] pListItem The item that the user wants to drag.
	/// \param[in] pListSubItem The sub-item that the user wants to drag. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemBeginDrag, ExplorerListView::Raise_ItemBeginDrag,
	///       Fire_ItemBeginRDrag, Fire_ColumnBeginDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemBeginDrag, ExplorerListView::Raise_ItemBeginDrag,
	///       Fire_ItemBeginRDrag, Fire_ColumnBeginDrag
	/// \endif
	HRESULT Fire_ItemBeginDrag(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemBeginRDrag event</em>
	///
	/// \param[in] pListItem The item that the user wants to drag.
	/// \param[in] pListSubItem The sub-item that the user wants to drag. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemBeginRDrag, ExplorerListView::Raise_ItemBeginRDrag,
	///       Fire_ItemBeginDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemBeginRDrag, ExplorerListView::Raise_ItemBeginRDrag,
	///       Fire_ItemBeginDrag
	/// \endif
	HRESULT Fire_ItemBeginRDrag(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetDisplayInfo event</em>
	///
	/// \param[in] pListItem The item whose properties are requested.
	/// \param[in] pListSubItem The sub-item whose properties are requested. May be \c NULL.
	/// \param[in] requestedInfo Specifies which properties' values are required. Any combination of
	///            the values defined by the \c RequestedInfoConstants enumeration is valid.
	/// \param[out] pIconIndex The zero-based index of the item's or sub-item's icon. If the \c requestedInfo
	///             parameter doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[out] pIndent The requested indentation. If the \c requestedInfo parameter doesn't
	///             include \c riIndent, the caller should ignore this value.
	/// \param[out] pGroupID The unique ID of the item's listview group. If the \c requestedInfo parameter
	///             doesn't include \c riGroupID, the caller should ignore this value.
	/// \param[out] ppTileViewColumns The requested array of \c TILEVIEWSUBITEM structs. If the
	///             \c requestedInfo parameter doesn't include \c riTileViewColumns, the caller should ignore
	///             this value.
	/// \param[in] maxItemTextLength The maximum number of characters the item's or sub-item's text may
	///            consist of. If the \c requestedInfo parameter doesn't include \c riItemText, the client
	///            should ignore this value.
	/// \param[out] pItemText The item's or sub-item's text. If the \c requestedInfo parameter doesn't include
	///             \c riItemText, the caller should ignore this value.
	/// \param[out] pOverlayIndex The zero-based index of the item's or sub-item's overlay icon. If the
	///             \c requestedInfo parameter doesn't include \c riOverlayIndex, the caller should ignore
	///             this value.
	/// \param[out] pStateImageIndex The one-based index of the item's or sub-item's state image. If the
	///             \c requestedInfo parameter doesn't include \c riStateImageIndex, the caller should
	///             ignore this value.
	/// \param[out] pItemStates A bit field specifying which of the following item or sub-item properties are
	///             equal to \c VARIANT_TRUE: \n
	///             - \c ListViewItem::get_Activating, ListViewSubItem::get_Activating
	///             - \c ListViewItem::get_Caret
	///             - \c ListViewItem::get_DropHilited
	///             - \c ListViewItem::get_Ghosted, ListViewSubItem::get_Ghosted
	///             - \c ListViewItem::get_Glowing, ListViewSubItem::get_Glowing
	///             - \c ListViewItem::get_Selected
	///
	///             Any combination of the following values defined by the \c ItemStateConstants
	///             enumeration is valid:\n
	///             - \c isActivating If the \c requestedInfo parameter doesn't include \c riActivating, the
	///               caller should ignore this flag.
	///             - \c isCaret If the \c requestedInfo parameter doesn't include \c riCaret, the caller
	///               should ignore this flag.
	///             - \c isDropHilited If the \c requestedInfo parameter doesn't include \c riDropHilited,
	///               the caller should ignore this flag.
	///             - \c isGhosted If the \c requestedInfo parameter doesn't include \c riGhosted, the caller
	///               should ignore this flag.
	///             - \c isGlowing If the \c requestedInfo parameter doesn't include \c riGlowing, the caller
	///               should ignore this flag.
	///             - \c isSelected If the \c requestedInfo parameter doesn't include \c riSelected, the
	///               caller should ignore this flag.
	/// \param[in,out] pDontAskAgain If \c VARIANT_TRUE, the caller should always use the same settings
	///                and never fire this event again for these properties of this item or sub-item;
	///                otherwise it shouldn't make the values persistent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemGetDisplayInfo,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, Fire_ItemSetText
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemGetDisplayInfo,
	///       ExplorerListView::Raise_ItemGetDisplayInfo, Fire_ItemSetText
	/// \endif
	HRESULT Fire_ItemGetDisplayInfo(IListViewItem* pListItem, IListViewSubItem* pListSubItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pIndent, LONG* pGroupID, LPSAFEARRAY* ppTileViewColumns, LONG maxItemTextLength, BSTR* pItemText, LONG* pOverlayIndex, LONG* pStateImageIndex, ItemStateConstants* pItemStates, VARIANT_BOOL* pDontAskAgain)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[13];
				p[12] = pListItem;
				p[11] = pListSubItem;
				p[10] = static_cast<LONG>(requestedInfo);							p[10].vt = VT_I4;
				p[9].plVal = pIconIndex;															p[9].vt = VT_I4 | VT_BYREF;
				p[8].plVal = pIndent;																	p[8].vt = VT_I4 | VT_BYREF;
				p[7].plVal = pGroupID;																p[7].vt = VT_I4 | VT_BYREF;
				p[6].pparray = ppTileViewColumns;											p[6].vt = VT_ARRAY | VT_RECORD | VT_BYREF;
				p[5] = maxItemTextLength;
				p[4].pbstrVal = pItemText;														p[4].vt = VT_BSTR | VT_BYREF;
				p[3].plVal = pOverlayIndex;														p[3].vt = VT_I4 | VT_BYREF;
				p[2].plVal = pStateImageIndex;												p[2].vt = VT_I4 | VT_BYREF;
				p[1].plVal = reinterpret_cast<PLONG>(pItemStates);		p[1].vt = VT_I4 | VT_BYREF;
				p[0].pboolVal = pDontAskAgain;												p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 13, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMGETDISPLAYINFO, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetGroup event</em>
	///
	/// \param[in] itemIndex The item's zero-based (control-wide) index.
	/// \param[in] occurrenceIndex The zero-based index of the item's copy for which the group membership is
	///            retrieved.
	/// \param[out] pGroupIndex Must be set to the zero-based index of the listview group that shall contain
	///             the specified copy of the specified item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemGetGroup, ExplorerListView::Raise_ItemGetGroup,
	///       Fire_ItemGetOccurrencesCount
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemGetGroup, ExplorerListView::Raise_ItemGetGroup,
	///       Fire_ItemGetOccurrencesCount
	/// \endif
	HRESULT Fire_ItemGetGroup(LONG itemIndex, LONG occurrenceIndex, LONG* pGroupIndex)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = itemIndex;
				p[1] = occurrenceIndex;
				p[0].plVal = pGroupIndex;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMGETGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetInfoTipText event</em>
	///
	/// \param[in] pListItem The item whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The item's info tip text.
	/// \param[in,out] pAbortToolTip If \c VARIANT_TRUE, the caller should NOT show the tooltip;
	///                otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemGetInfoTipText,
	///       ExplorerListView::Raise_ItemGetInfoTipText, Fire_SettingItemInfoTipText
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemGetInfoTipText,
	///       ExplorerListView::Raise_ItemGetInfoTipText, Fire_SettingItemInfoTipText
	/// \endif
	HRESULT Fire_ItemGetInfoTipText(IListViewItem* pListItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListItem;
				p[2] = maxInfoTipLength;
				p[1].pbstrVal = pInfoTipText;			p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pAbortToolTip;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMGETINFOTIPTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemGetOccurrencesCount event</em>
	///
	/// \param[in] itemIndex The item's zero-based (control-wide) index.
	/// \param[out] pOccurrencesCount Must be set to the number of occurrences of the item in the listview
	///             control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemGetOccurrencesCount,
	///       ExplorerListView::Raise_ItemGetOccurrencesCount, Fire_ItemGetGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemGetOccurrencesCount,
	///       ExplorerListView::Raise_ItemGetOccurrencesCount, Fire_ItemGetGroup
	/// \endif
	HRESULT Fire_ItemGetOccurrencesCount(LONG itemIndex, LONG* pOccurrencesCount)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = itemIndex;
				p[0].plVal = pOccurrencesCount;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMGETOCCURRENCESCOUNT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseEnter event</em>
	///
	/// \param[in] pListItem The item that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemMouseEnter, ExplorerListView::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_SubItemMouseEnter, Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemMouseEnter, ExplorerListView::Raise_ItemMouseEnter,
	///       Fire_ItemMouseLeave, Fire_SubItemMouseEnter, Fire_MouseMove
	/// \endif
	HRESULT Fire_ItemMouseEnter(IListViewItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemMouseLeave event</em>
	///
	/// \param[in] pListItem The item that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemMouseLeave, ExplorerListView::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_SubItemMouseLeave, Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemMouseLeave, ExplorerListView::Raise_ItemMouseLeave,
	///       Fire_ItemMouseEnter, Fire_SubItemMouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_ItemMouseLeave(IListViewItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemSelectionChanged event</em>
	///
	/// \param[in] pListItem The item that was selected/deselected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemSelectionChanged,
	///       ExplorerListView::Raise_ItemSelectionChanged, Fire_BeginMarqueeSelection,
	///       Fire_SelectedItemRange, Fire_CaretChanged, Fire_GroupSelectionChanged
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemSelectionChanged,
	///       ExplorerListView::Raise_ItemSelectionChanged, Fire_BeginMarqueeSelection,
	///       Fire_SelectedItemRange, Fire_CaretChanged, Fire_GroupSelectionChanged
	/// \endif
	HRESULT Fire_ItemSelectionChanged(IListViewItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMSELECTIONCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemSetText event</em>
	///
	/// \param[in] pListItem The item whose text was changed.
	/// \param[in] itemText The item's new text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemSetText, ExplorerListView::Raise_ItemSetText,
	///       Fire_ItemGetDisplayInfo
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemSetText, ExplorerListView::Raise_ItemSetText,
	///       Fire_ItemGetDisplayInfo
	/// \endif
	HRESULT Fire_ItemSetText(IListViewItem* pListItem, BSTR itemText)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0] = itemText;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMSETTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemStateImageChanged event</em>
	///
	/// \param[in] pListItem The item whose state image was changed.
	/// \param[in] previousStateImageIndex The one-based index of the item's previous state image.
	/// \param[in] newStateImageIndex The one-based index of the item's new state image.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemStateImageChanged,
	///       ExplorerListView::Raise_ItemStateImageChanged, Fire_ItemStateImageChanging,
	///       Fire_ColumnStateImageChanged
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemStateImageChanged,
	///       ExplorerListView::Raise_ItemStateImageChanged, Fire_ItemStateImageChanging,
	///       Fire_ColumnStateImageChanged
	/// \endif
	HRESULT Fire_ItemStateImageChanged(IListViewItem* pListItem, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListItem;
				p[2] = previousStateImageIndex;
				p[1] = newStateImageIndex;
				p[0] = static_cast<LONG>(causedBy);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMSTATEIMAGECHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ItemStateImageChanging event</em>
	///
	/// \param[in] pListItem The item whose state image shall be changed.
	/// \param[in] previousStateImageIndex The one-based index of the item's previous state image.
	/// \param[in,out] pNewStateImageIndex The one-based index of the item's new state image. The client
	///                may change this value.
	/// \param[in] causedBy The reason for the state image change. Any of the values defined by the
	///            \c StateImageChangeCausedByConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should prevent the state image from
	///                changing; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ItemStateImageChanging,
	///       ExplorerListView::Raise_ItemStateImageChanging, Fire_ItemStateImageChanged,
	///       Fire_ColumnStateImageChanging
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ItemStateImageChanging,
	///       ExplorerListView::Raise_ItemStateImageChanging, Fire_ItemStateImageChanged,
	///       Fire_ColumnStateImageChanging
	/// \endif
	HRESULT Fire_ItemStateImageChanging(IListViewItem* pListItem, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pListItem;
				p[3] = previousStateImageIndex;
				p[2].plVal = pNewStateImageIndex;			p[2].vt = VT_I4 | VT_BYREF;
				p[1] = static_cast<LONG>(causedBy);		p[1].vt = VT_I4;
				p[0].pboolVal = pCancelChange;				p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_ITEMSTATEIMAGECHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::KeyDown, ExplorerListView::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::KeyDown, ExplorerListView::Raise_KeyDown, Fire_KeyUp,
	///       Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::KeyPress, ExplorerListView::Raise_KeyPress,
	///       Fire_IncrementalSearchStringChanging, Fire_KeyDown, Fire_KeyUp
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::KeyPress, ExplorerListView::Raise_KeyPress,
	///       Fire_IncrementalSearchStringChanging, Fire_KeyDown, Fire_KeyUp
	/// \endif
	HRESULT Fire_KeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::KeyUp, ExplorerListView::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::KeyUp, ExplorerListView::Raise_KeyUp, Fire_KeyDown,
	///       Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MapGroupWideToTotalItemIndex event</em>
	///
	/// \param[in] groupIndex The zero-based index of the item's listview group.
	/// \param[in] groupWideItemIndex The item's zero-based group-wide index.
	/// \param[out] pTotalItemIndex Must be set to the item's zero-based total index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MapGroupWideToTotalItemIndex,
	///       ExplorerListView::Raise_MapGroupWideToTotalItemIndex, Fire_ItemGetDisplayInfo
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MapGroupWideToTotalItemIndex,
	///       ExplorerListView::Raise_MapGroupWideToTotalItemIndex, Fire_ItemGetDisplayInfo
	/// \endif
	HRESULT Fire_MapGroupWideToTotalItemIndex(LONG groupIndex, LONG groupWideItemIndex, LONG* pTotalItemIndex)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = groupIndex;
				p[1] = groupWideItemIndex;
				p[0].plVal = pTotalItemIndex;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MAPGROUPWIDETOTOTALITEMINDEX, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
	/// \param[in] pListSubItem The clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MClick, ExplorerListView::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MClick, ExplorerListView::Raise_MClick, Fire_MDblClick,
	///       Fire_Click, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
	/// \param[in] pListSubItem The double-clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MDblClick, ExplorerListView::Raise_MDblClick,
	///       Fire_MClick, Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MDblClick, ExplorerListView::Raise_MDblClick,
	///       Fire_MClick, Fire_DblClick, Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseDown, ExplorerListView::Raise_MouseDown,
	///       Fire_MouseUp, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseDown, ExplorerListView::Raise_MouseDown,
	///       Fire_MouseUp, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseEnter, ExplorerListView::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_ItemMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseEnter, ExplorerListView::Raise_MouseEnter,
	///       Fire_MouseLeave, Fire_ItemMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseHover, ExplorerListView::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseHover, ExplorerListView::Raise_MouseHover,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseLeave, ExplorerListView::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_ItemMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseLeave, ExplorerListView::Raise_MouseLeave,
	///       Fire_MouseEnter, Fire_ItemMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseMove, ExplorerListView::Raise_MouseMove,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseMove, ExplorerListView::Raise_MouseMove,
	///       Fire_MouseEnter, Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp, Fire_MouseWheel
	/// \endif
	HRESULT Fire_MouseMove(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseUp, ExplorerListView::Raise_MouseUp,
	///       Fire_MouseDown, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseUp, ExplorerListView::Raise_MouseUp,
	///       Fire_MouseDown, Fire_Click, Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseWheel event</em>
	///
	/// \param[in] pListItem The item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] pListSubItem The sub-item that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in] scrollAxis Specifies whether the user intents to scroll vertically or horizontally.
	///            Any of the values defined by the \c ScrollAxisConstants enumeration is valid.
	/// \param[in] wheelDelta The distance the wheel has been rotated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::MouseWheel, ExplorerListView::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_EditMouseWheel, Fire_HeaderMouseWheel
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::MouseWheel, ExplorerListView::Raise_MouseWheel,
	///       Fire_MouseMove, Fire_EditMouseWheel, Fire_HeaderMouseWheel
	/// \endif
	HRESULT Fire_MouseWheel(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pListItem;
				p[7] = pListSubItem;
				p[6] = button;																		p[6].vt = VT_I2;
				p[5] = shift;																			p[5].vt = VT_I2;
				p[4] = x;																					p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																					p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);		p[2].vt = VT_I4;
				p[1].lVal = static_cast<LONG>(scrollAxis);				p[1].vt = VT_I4;
				p[0] = wheelDelta;																p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_MOUSEWHEEL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLECompleteDrag, ExplorerListView::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag, Fire_HeaderOLECompleteDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLECompleteDrag, ExplorerListView::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag, Fire_HeaderOLECompleteDrag
	/// \endif
	HRESULT Fire_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pData;
				p[0].lVal = static_cast<LONG>(performedEffect);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in,out] ppDropTarget The item that is the target of the drag'n'drop operation. The client
	///                may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEDragDrop, ExplorerListView::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp,
	///       Fire_HeaderOLEDragDrop
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEDragDrop, ExplorerListView::Raise_OLEDragDrop,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp,
	///       Fire_HeaderOLEDragDrop
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);									p[6].vt = VT_I4 | VT_BYREF;
				p[5].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[5].vt = VT_DISPATCH | VT_BYREF;
				p[4] = button;																									p[4].vt = VT_I2;
				p[3] = shift;																										p[3].vt = VT_I2;
				p[2] = x;																												p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																												p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);									p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEDragEnter, ExplorerListView::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_HeaderOLEDragEnter
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEDragEnter, ExplorerListView::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter,
	///       Fire_HeaderOLEDragEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pData;
				p[8].plVal = reinterpret_cast<PLONG>(pEffect);									p[8].vt = VT_I4 | VT_BYREF;
				p[7].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[7].vt = VT_DISPATCH | VT_BYREF;
				p[6] = button;																									p[6].vt = VT_I2;
				p[5] = shift;																										p[5].vt = VT_I2;
				p[4] = x;																												p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																												p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The handle of the potential drag'n'drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEDragEnterPotentialTarget,
	///       ExplorerListView::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget,
	///       Fire_HeaderOLEDragEnterPotentialTarget
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEDragEnterPotentialTarget,
	///       ExplorerListView::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget,
	///       Fire_HeaderOLEDragEnterPotentialTarget
	/// \endif
	HRESULT Fire_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndPotentialTarget;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The item that is the current target of the drag'n'drop operation.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEDragLeave, ExplorerListView::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave,
	///       Fire_HeaderOLEDragEnter
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEDragLeave, ExplorerListView::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave,
	///       Fire_HeaderOLEDragEnter
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, IListViewItem* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pData;
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEDragLeavePotentialTarget,
	///       ExplorerListView::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget,
	///       Fire_HeaderOLEDragLeavePotentialTarget
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEDragLeavePotentialTarget,
	///       ExplorerListView::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget,
	///       Fire_HeaderOLEDragLeavePotentialTarget
	/// \endif
	HRESULT Fire_OLEDragLeavePotentialTarget(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The item that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another item.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[in,out] pAutoHScrollVelocity The speed multiplier for horizontal auto-scrolling. If set to 0,
	///                the caller should disable horizontal auto-scrolling; if set to a value less than 0, it
	///                should scroll the control to the left; if set to a value greater than 0, it should
	///                scroll the control to the right. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	/// \param[in,out] pAutoVScrollVelocity The speed multiplier for vertical auto-scrolling. If set to 0,
	///                the caller should disable vertical auto-scrolling; if set to a value less than 0, it
	///                should scroll the control upwardly; if set to a value greater than 0, it should
	///                scroll the control downwards. The higher/lower the value is, the faster the caller
	///                should scroll the control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEDragMouseMove,
	///       ExplorerListView::Raise_OLEDragMouseMove, Fire_OLEDragEnter, Fire_OLEDragLeave,
	///       Fire_OLEDragDrop, Fire_MouseMove, Fire_HeaderOLEDragMouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEDragMouseMove,
	///       ExplorerListView::Raise_OLEDragMouseMove, Fire_OLEDragEnter, Fire_OLEDragLeave,
	///       Fire_OLEDragDrop, Fire_MouseMove, Fire_HeaderOLEDragMouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IListViewItem** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pAutoHScrollVelocity, LONG* pAutoVScrollVelocity)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[10];
				p[9] = pData;
				p[8].plVal = reinterpret_cast<PLONG>(pEffect);									p[8].vt = VT_I4 | VT_BYREF;
				p[7].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[7].vt = VT_DISPATCH | VT_BYREF;
				p[6] = button;																									p[6].vt = VT_I2;
				p[5] = shift;																										p[5].vt = VT_I2;
				p[4] = x;																												p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;																												p[3].vt = VT_YPOS_PIXELS;
				p[2].lVal = static_cast<LONG>(hitTestDetails);									p[2].vt = VT_I4;
				p[1].plVal = pAutoHScrollVelocity;															p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = pAutoVScrollVelocity;															p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 10, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c OLEDropEffectConstants enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEGiveFeedback,
	///       ExplorerListView::Raise_OLEGiveFeedback, Fire_OLEQueryContinueDrag, Fire_HeaderOLEGiveFeedback
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEGiveFeedback,
	///       ExplorerListView::Raise_OLEGiveFeedback, Fire_OLEQueryContinueDrag, Fire_HeaderOLEGiveFeedback
	/// \endif
	HRESULT Fire_OLEGiveFeedback(OLEDropEffectConstants effect, VARIANT_BOOL* pUseDefaultCursors)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = static_cast<LONG>(effect);			p[1].vt = VT_I4;
				p[0].pboolVal = pUseDefaultCursors;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c VARIANT_TRUE, the user has pressed the \c ESC key since the last
	///            time this event was fired; otherwise not.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue, cancel or complete the
	///                drag'n'drop operation. Any of the values defined by the
	///                \c OLEActionToContinueWithConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEQueryContinueDrag,
	///       ExplorerListView::Raise_OLEQueryContinueDrag, Fire_OLEGiveFeedback,
	///       Fire_HeaderOLEQueryContinueDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEQueryContinueDrag,
	///       ExplorerListView::Raise_OLEQueryContinueDrag, Fire_OLEGiveFeedback,
	///       Fire_HeaderOLEQueryContinueDrag
	/// \endif
	HRESULT Fire_OLEQueryContinueDrag(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, OLEActionToContinueWithConstants* pActionToContinueWith)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pressedEscape;
				p[2] = button;																									p[2].vt = VT_I2;
				p[1] = shift;																										p[1].vt = VT_I2;
				p[0].plVal = reinterpret_cast<PLONG>(pActionToContinueWith);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEReceivedNewData,
	///       ExplorerListView::Raise_OLEReceivedNewData, Fire_HeaderOLEReceivedNewData, Fire_OLESetData
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEReceivedNewData,
	///       ExplorerListView::Raise_OLEReceivedNewData, Fire_HeaderOLEReceivedNewData, Fire_OLESetData
	/// \endif
	HRESULT Fire_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLESetData,
	///       ExplorerListView::Raise_OLESetData, Fire_OLEStartDrag, Fire_HeaderOLESetData
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLESetData,
	///       ExplorerListView::Raise_OLESetData, Fire_OLEStartDrag, Fire_HeaderOLESetData
	/// \endif
	HRESULT Fire_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OLEStartDrag, ExplorerListView::Raise_OLEStartDrag,
	///       Fire_OLESetData, Fire_OLECompleteDrag, Fire_HeaderOLEStartDrag
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OLEStartDrag, ExplorerListView::Raise_OLEStartDrag,
	///       Fire_OLESetData, Fire_OLECompleteDrag, Fire_HeaderOLEStartDrag
	/// \endif
	HRESULT Fire_OLEStartDrag(IOLEDataObject* pData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OwnerDrawItem event</em>
	///
	/// \param[in] pListItem The item to draw.
	/// \param[in] itemState The item's current state (focused, selected etc.). Most of the values
	///            defined by the \c OwnerDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::OwnerDrawItem, ExplorerListView::Raise_OwnerDrawItem,
	///       Fire_CustomDraw, Fire_HeaderOwnerDrawItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::OwnerDrawItem, ExplorerListView::Raise_OwnerDrawItem,
	///       Fire_CustomDraw, Fire_HeaderOwnerDrawItem
	/// \endif
	HRESULT Fire_OwnerDrawItem(IListViewItem* pListItem, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListItem;
				p[2].lVal = static_cast<LONG>(itemState);		p[2].vt = VT_I4;
				p[1] = hDC;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{4C7D8A62-0F3A-41a7-BD58-FF25497D8451}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{00B2ADA1-29FD-45ce-A4C3-A2B85928FADD}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[0]);
				p[0].vt = VT_RECORD | VT_BYREF;
				p[0].pRecInfo = pRecordInfo;
				p[0].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_OWNERDRAWITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[0].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
	/// \param[in] pListSubItem The clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RClick, ExplorerListView::Raise_RClick,
	///       Fire_ContextMenu, Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RClick, ExplorerListView::Raise_RClick,
	///       Fire_ContextMenu, Fire_RDblClick, Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
	/// \param[in] pListSubItem The double-clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RDblClick, ExplorerListView::Raise_RDblClick,
	///       Fire_RClick, Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RDblClick, ExplorerListView::Raise_RDblClick,
	///       Fire_RClick, Fire_DblClick, Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RecreatedControlWindow,
	///       ExplorerListView::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RecreatedControlWindow,
	///       ExplorerListView::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \endif
	HRESULT Fire_RecreatedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedColumn event</em>
	///
	/// \param[in] pColumn The removed column.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RemovedColumn, ExplorerListView::Raise_RemovedColumn,
	///       Fire_RemovingColumn, Fire_InsertedColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RemovedColumn, ExplorerListView::Raise_RemovedColumn,
	///       Fire_RemovingColumn, Fire_InsertedColumn
	/// \endif
	HRESULT Fire_RemovedColumn(IVirtualListViewColumn* pColumn)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pColumn;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_REMOVEDCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedGroup event</em>
	///
	/// \param[in] pGroup The removed group.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RemovedGroup, ExplorerListView::Raise_RemovedGroup,
	///       Fire_RemovingGroup, Fire_InsertedGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RemovedGroup, ExplorerListView::Raise_RemovedGroup,
	///       Fire_RemovingGroup, Fire_InsertedGroup
	/// \endif
	HRESULT Fire_RemovedGroup(IVirtualListViewGroup* pGroup)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pGroup;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_REMOVEDGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedItem event</em>
	///
	/// \param[in] pListItem The removed item.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RemovedItem, ExplorerListView::Raise_RemovedItem,
	///       Fire_RemovingItem, Fire_InsertedItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RemovedItem, ExplorerListView::Raise_RemovedItem,
	///       Fire_RemovingItem, Fire_InsertedItem
	/// \endif
	HRESULT Fire_RemovedItem(IVirtualListViewItem* pListItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pListItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_REMOVEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingColumn event</em>
	///
	/// \param[in] pColumn The column being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RemovingColumn, ExplorerListView::Raise_RemovingColumn,
	///       Fire_RemovedColumn, Fire_InsertingColumn
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RemovingColumn, ExplorerListView::Raise_RemovingColumn,
	///       Fire_RemovedColumn, Fire_InsertingColumn
	/// \endif
	HRESULT Fire_RemovingColumn(IListViewColumn* pColumn, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pColumn;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_REMOVINGCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingGroup event</em>
	///
	/// \param[in] pGroup The group being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RemovingGroup, ExplorerListView::Raise_RemovingGroup,
	///       Fire_RemovedGroup, Fire_InsertingGroup
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RemovingGroup, ExplorerListView::Raise_RemovingGroup,
	///       Fire_RemovedGroup, Fire_InsertingGroup
	/// \endif
	HRESULT Fire_RemovingGroup(IListViewGroup* pGroup, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pGroup;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_REMOVINGGROUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingItem event</em>
	///
	/// \param[in] pListItem The item being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RemovingItem, ExplorerListView::Raise_RemovingItem,
	///       Fire_RemovedItem, Fire_InsertingItem
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RemovingItem, ExplorerListView::Raise_RemovingItem,
	///       Fire_RemovedItem, Fire_InsertingItem
	/// \endif
	HRESULT Fire_RemovingItem(IListViewItem* pListItem, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_REMOVINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RenamedItem event</em>
	///
	/// \param[in] pListItem The renamed item.
	/// \param[in] previousItemText The item's previous text.
	/// \param[in] newItemText The item's new text.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RenamedItem, ExplorerListView::Raise_RenamedItem,
	///       Fire_RenamingItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RenamedItem, ExplorerListView::Raise_RenamedItem,
	///       Fire_RenamingItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \endif
	HRESULT Fire_RenamedItem(IListViewItem* pListItem, BSTR previousItemText, BSTR newItemText)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pListItem;
				p[1] = previousItemText;
				p[0] = newItemText;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RENAMEDITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RenamingItem event</em>
	///
	/// \param[in] pListItem The item being renamed.
	/// \param[in] previousItemText The item's previous text.
	/// \param[in] newItemText The item's new text.
	/// \param[in,out] pCancelRenaming If \c VARIANT_TRUE, the caller should abort renaming; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::RenamingItem, ExplorerListView::Raise_RenamingItem,
	///       Fire_RenamedItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::RenamingItem, ExplorerListView::Raise_RenamingItem,
	///       Fire_RenamedItem, Fire_StartingLabelEditing, Fire_DestroyedEditControlWindow
	/// \endif
	HRESULT Fire_RenamingItem(IListViewItem* pListItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* pCancelRenaming)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pListItem;
				p[2] = previousItemText;
				p[1] = newItemText;
				p[0].pboolVal = pCancelRenaming;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RENAMINGITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ResizedControlWindow,
	///       ExplorerListView::Raise_ResizedControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ResizedControlWindow,
	///       ExplorerListView::Raise_ResizedControlWindow
	/// \endif
	HRESULT Fire_ResizedControlWindow(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizingColumn event</em>
	///
	/// \param[in] pColumn The column being resized.
	/// \param[in] pNewColumnWidth The column's new width in pixels. The client may change this value.
	/// \param[in,out] pAbortResizing If \c VARIANT_TRUE, the caller should abort column resizing; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::ResizingColumn,
	///       ExplorerListView::Raise_ResizingColumn, Fire_BeginColumnResizing, Fire_EndColumnResizing
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::ResizingColumn,
	///       ExplorerListView::Raise_ResizingColumn, Fire_BeginColumnResizing, Fire_EndColumnResizing
	/// \endif
	HRESULT Fire_ResizingColumn(IListViewColumn* pColumn, LONG* pNewColumnWidth, VARIANT_BOOL* pAbortResizing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pColumn;
				p[1].plVal = static_cast<PLONG>(pNewColumnWidth);		p[1].vt = VT_I4 | VT_BYREF;
				p[0].pboolVal = pAbortResizing;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_RESIZINGCOLUMN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectedItemRange event</em>
	///
	/// \param[in] pFirstItem The first item that was selected.
	/// \param[in] pLastItem The last item that was selected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::SelectedItemRange,
	///       ExplorerListView::Raise_SelectedItemRange, Fire_ItemSelectionChanged,
	///       Fire_BeginMarqueeSelection
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::SelectedItemRange,
	///       ExplorerListView::Raise_SelectedItemRange, Fire_ItemSelectionChanged,
	///       Fire_BeginMarqueeSelection
	/// \endif
	HRESULT Fire_SelectedItemRange(IListViewItem* pFirstItem, IListViewItem* pLastItem)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pFirstItem;
				p[0] = pLastItem;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_SELECTEDITEMRANGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SettingItemInfoTipText event</em>
	///
	/// \param[in] pListItem The item for which the info tip text is being set.
	/// \param[in,out] pInfoTipText The info tip text being set. It may be changed by the client application.
	/// \param[in,out] pAbortInfoTip If \c VARIANT_TRUE, the caller should cancel setting the info tip;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::SettingItemInfoTipText,
	///       ExplorerListView::Raise_SettingItemInfoTipText, Fire_ItemGetInfoTipText
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::SettingItemInfoTipText,
	///       ExplorerListView::Raise_SettingItemInfoTipText, Fire_ItemGetInfoTipText
	/// \endif
	HRESULT Fire_SettingItemInfoTipText(IListViewItem* pListItem, BSTR* pInfoTipText, VARIANT_BOOL* pAbortInfoTip)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pListItem;
				p[1].pbstrVal = pInfoTipText;			p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pAbortInfoTip;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_SETTINGITEMINFOTIPTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c StartingLabelEditing event</em>
	///
	/// \param[in] pListItem The item being edited.
	/// \param[in,out] pCancelEditing If \c VARIANT_TRUE, the caller should prevent label-editing;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::StartingLabelEditing,
	///       ExplorerListView::Raise_StartingLabelEditing, Fire_RenamingItem, Fire_CreatedEditControlWindow
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::StartingLabelEditing,
	///       ExplorerListView::Raise_StartingLabelEditing, Fire_RenamingItem, Fire_CreatedEditControlWindow
	/// \endif
	HRESULT Fire_StartingLabelEditing(IListViewItem* pListItem, VARIANT_BOOL* pCancelEditing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pListItem;
				p[0].pboolVal = pCancelEditing;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_STARTINGLABELEDITING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SubItemMouseEnter event</em>
	///
	/// \param[in] pListSubItem The sub-item that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::SubItemMouseEnter,
	///       ExplorerListView::Raise_SubItemMouseEnter, Fire_SubItemMouseLeave, Fire_ItemMouseEnter,
	///       Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::SubItemMouseEnter,
	///       ExplorerListView::Raise_SubItemMouseEnter, Fire_SubItemMouseLeave, Fire_ItemMouseEnter,
	///       Fire_MouseMove
	/// \endif
	HRESULT Fire_SubItemMouseEnter(IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_SUBITEMMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SubItemMouseLeave event</em>
	///
	/// \param[in] pListSubItem The sub-item that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::SubItemMouseLeave,
	///       ExplorerListView::Raise_SubItemMouseLeave, Fire_SubItemMouseEnter, Fire_ItemMouseLeave,
	///       Fire_MouseMove
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::SubItemMouseLeave,
	///       ExplorerListView::Raise_SubItemMouseLeave, Fire_SubItemMouseEnter, Fire_ItemMouseLeave,
	///       Fire_MouseMove
	/// \endif
	HRESULT Fire_SubItemMouseLeave(IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_SUBITEMMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] pListItem The clicked item. May be \c NULL.
	/// \param[in] pListSubItem The clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be a
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::XClick, ExplorerListView::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::XClick, ExplorerListView::Raise_XClick, Fire_XDblClick,
	///       Fire_Click, Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] pListItem The double-clicked item. May be \c NULL.
	/// \param[in] pListSubItem The double-clicked sub-item. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa ExLVwLibU::_IExplorerListViewEvents::XDblClick, ExplorerListView::Raise_XDblClick,
	///       Fire_XClick, Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa ExLVwLibA::_IExplorerListViewEvents::XDblClick, ExplorerListView::Raise_XDblClick,
	///       Fire_XClick, Fire_DblClick, Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pListItem;
				p[5] = pListSubItem;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_EXLVWE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IExplorerListViewEvents