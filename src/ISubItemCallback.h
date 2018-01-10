//////////////////////////////////////////////////////////////////////
/// \class ISubItemCallback
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for supporting enhanced label-editing and enhanced owner-drawing of sub-items</em>
///
/// This interface is implemented by client applications and used by listview controls (comctl32.dll
/// version 6.10 or higher) to provide enhanced label-editing features. By implementing this interface,
/// the following features become possible:
/// - edit listview group header labels
/// - edit sub-item labels
/// - provide other control's than a text box for label-editing (it should also be possible to edit labels
///   without any control)
/// - start label-editing mode on mouse hover\n
/// Also this interface provides methods to simplify owner-drawing of single sub-items.
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \todo Improve documentation.
///
/// \sa ExplorerListView, IListView_WINVISTA, IListView_WIN7, IPropertyControl, IDrawPropertyControl,
///     <a href="https://www.geoffchappell.com/studies/windows/shell/comctl32/controls/listview/interfaces/isubitemcallback.htm">ISubItemCallback</a>
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IPropertyControl.h"


// the interface's GUID
const IID IID_ISubItemCallback = {0x11A66240, 0x5489, 0x42C2, {0xAE, 0xBF, 0x28, 0x6F, 0xC8, 0x31, 0x52, 0x4C}};


class ISubItemCallback :
    public IUnknown
{
public:
	/// \brief <em>Retrieves the title of the specified sub-item</em>
	///
	/// Retrieves the title of the specified sub-item. This title is displayed in extended tile view mode
	/// in front of the sub-item's value.
	///
	/// \param[in] subItemIndex The one-based index of the sub-item for which to retrieve the title.
	/// \param[out] pBuffer Receives the sub-item's title.
	/// \param[in] bufferSize The size of the buffer (in characters) specified by the \c pBuffer parameter.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE GetSubItemTitle(int subItemIndex, LPWSTR pBuffer, int bufferSize) = 0;
	/// \brief <em>Retrieves the control representing the specified sub-item</em>
	///
	/// Retrieves the control representing the specified sub-item. The control is used for drawing the
	/// sub-item.\n
	/// Starting with comctl32.dll version 6.10, sub-items can be represented by objects that implement
	/// the \c IPropertyControlBase interface. Representation means drawing the sub-item (by implementing the
	/// \c IDrawPropertyControl interface) and/or editing the sub-item in-place (by implementing the
	/// \c IPropertyControl interface, which allows in-place editing with a complex user interface). The
	/// object that represents the sub-item is retrieved dynamically.\n
	/// This method retrieves the sub-item control that is used for drawing the sub-item.
	///
	/// \param[in] itemIndex The zero-based index of the item for which to retrieve the sub-item control.
	/// \param[in] subItemIndex The one-based index of the sub-item for which to retrieve the sub-item
	///            control.
	/// \param[in] requiredInterface The IID of the interface of which the sub-item control's implementation
	///            will be returned.
	/// \param[out] ppObject Receives the sub-item control's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks With current versions of comctl32.dll, providing a sub-item control is the only way to
	///          custom-draw sub-items in Tiles view mode.
	///
	/// \sa BeginSubItemEdit, IPropertyControlBase, IPropertyControl, IDrawPropertyControl
	virtual HRESULT STDMETHODCALLTYPE GetSubItemControl(int itemIndex, int subItemIndex, REFIID requiredInterface, LPVOID* ppObject) = 0;
	/// \brief <em>Retrieves the control used to edit the specified sub-item</em>
	///
	/// Retrieves the control that will be used to edit the specified sub-item. The control is used for
	/// editing the sub-item.\n
	/// Starting with comctl32.dll version 6.10, sub-items can be represented by objects that implement
	/// the \c IPropertyControlBase interface. Representation means drawing the sub-item (by implementing the
	/// \c IDrawPropertyControl interface) and/or editing the sub-item in-place (by implementing the
	/// \c IPropertyControl interface, which allows in-place editing with a complex user interface). The
	/// object that represents the sub-item is retrieved dynamically.\n
	/// This method retrieves the sub-item control that is used for editing the sub-item.
	///
	/// \param[in] itemIndex The zero-based index of the item for which to retrieve the sub-item control.
	/// \param[in] subItemIndex The one-based index of the sub-item for which to retrieve the sub-item
	///            control.
	/// \param[in] mode If set to 0, the edit mode has been entered by moving the mouse over the sub-item.
	///            If set to 1, the edit mode has been entered by calling \c IListView::EditSubItem or by
	///            clicking on the sub-item.
	/// \param[in] requiredInterface The IID of the interface of which the sub-item control's implementation
	///            will be returned.
	/// \param[out] ppObject Receives the sub-item control's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EndSubItemEdit, GetSubItemControl, IListView_WINVISTA::EditSubItem, IListView_WIN7::EditSubItem,
	///     IPropertyControlBase, IPropertyControl, IDrawPropertyControl
	virtual HRESULT STDMETHODCALLTYPE BeginSubItemEdit(int itemIndex, int subItemIndex, int mode, REFIID requiredInterface, LPVOID* ppObject) = 0;
	/// \brief <em>Notifies the control that editing the specified sub-item has ended</em>
	///
	/// Notifies the control that editing the specified sub-item using the specified sub-item control has
	/// been finished.\n
	/// Starting with comctl32.dll version 6.10, sub-items can be represented by objects that implement
	/// the \c IPropertyControlBase interface. Representation means drawing the sub-item (by implementing the
	/// \c IDrawPropertyControl interface) and/or editing the sub-item in-place (by implementing the
	/// \c IPropertyControl interface, which allows in-place editing with a complex user interface). The
	/// object that represents the sub-item is retrieved dynamically.
	///
	/// \param[in] itemIndex The zero-based index of the item whose sub-item has been edited.
	/// \param[in] subItemIndex The one-based index of the sub-item that has been edited.
	/// \param[in] mode If set to 0, the edit mode has been entered by moving the mouse over the sub-item.
	///            If set to 1, the edit mode has been entered by calling \c IListView::EditSubItem or by
	///            clicking on the sub-item.
	/// \param[in] pPropertyControl The property control that has been used to edit the sub-item. This
	///            property control has to be destroyed by this method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Call \c IPropertyControl::IsModified to retrieve whether editing the sub-item has been
	///          completed or canceled.
	///
	/// \sa BeginSubItemEdit, IListView_WINVISTA::EditSubItem, IListView_WIN7::EditSubItem,
	///     IPropertyControlBase, IPropertyControl, IDrawPropertyControl, IPropertyControlBase::Destroy,
	///     IPropertyControl::IsModified
	virtual HRESULT STDMETHODCALLTYPE EndSubItemEdit(int itemIndex, int subItemIndex, int mode, IPropertyControl* pPropertyControl) = 0;
	/// \brief <em>Retrieves the control used to edit the specified group header</em>
	///
	/// Retrieves the control that will be used to edit the specified group header. The control is used for
	/// editing the group header.\n
	/// Starting with comctl32.dll version 6.10, group headers can be represented by objects that implement
	/// the \c IPropertyControlBase interface. Representation means editing the group header in-place (by
	/// implementing the \c IPropertyControl interface, which allows in-place editing with a complex user
	/// interface). The object that represents the group header is retrieved dynamically.\n
	/// This method retrieves the edit control that is used for editing the group header.
	///
	/// \param[in] groupIndex The zero-based index of the group for which to retrieve the edit control.
	/// \param[in] requiredInterface The IID of the interface of which the edit control's implementation will
	///            be returned.
	/// \param[out] ppObject Receives the edit control's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EndGroupEdit, IListView_WINVISTA::EditGroupLabel, IListView_WIN7::EditGroupLabel, IPropertyControlBase,
	///     IPropertyControl
	virtual HRESULT STDMETHODCALLTYPE BeginGroupEdit(int groupIndex, REFIID requiredInterface, LPVOID* ppObject) = 0;
	/// \brief <em>Notifies the control that editing the specified group header has ended</em>
	///
	/// Notifies the control that editing the specified group header using the specified edit control has
	/// been finished.\n
	/// Starting with comctl32.dll version 6.10, group headers can be represented by objects that implement
	/// the \c IPropertyControlBase interface. Representation means editing the group header in-place (by
	/// implementing the \c IPropertyControl interface, which allows in-place editing with a complex user
	/// interface). The object that represents the group header is retrieved dynamically.
	///
	/// \param[in] groupIndex The zero-based index of the group whose header has been edited.
	/// \param[in] mode If set to 0, the edit mode has been entered by moving the mouse over the group
	///            header. If set to 1, the edit mode has been entered by calling
	///            \c IListView::EditGroupLabel or by clicking on the group header.
	/// \param[in] pPropertyControl The property control that has been used to edit the group header. This
	///            property control has to be destroyed by this method.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Call \c IPropertyControl::IsModified to retrieve whether editing the group header has been
	///          completed or canceled.
	///
	/// \sa BeginGroupEdit, IListView_WINVISTA::EditGroupLabel, IListView_WIN7::EditGroupLabel,
	///     IPropertyControlBase, IPropertyControl, IPropertyControlBase::Destroy,
	///     IPropertyControl::IsModified
	virtual HRESULT STDMETHODCALLTYPE EndGroupEdit(int groupIndex, int mode, IPropertyControl* pPropertyControl) = 0;
	/// \brief <em>Notifies the control that the specified verb has been invoked on the specified item</em>
	///
	/// Notifies the control that an action identified by the specified verb has been invoked on the
	/// specified item. The action usually is generated by the user. The sub-item control translates the
	/// action into a verb which it invokes.\n
	/// For the hyperlink sub-item control the action is clicking the link and the verb is the string that
	/// has been specified as the \c id attribute of the hyperlink markup
	/// (&lt;a id=&quot;<i>verb</i>&quot;&gt;<i>text</i>&lt;/a&gt;).
	///
	/// \param[in] itemIndex The zero-based index of the item on which the verb is being invoked.
	/// \param[in] pVerb The verb identifying the action.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetSubItemControl, IPropertyControlBase::InvokeDefaultAction
	virtual HRESULT STDMETHODCALLTYPE OnInvokeVerb(int itemIndex, LPCWSTR pVerb) = 0;
};