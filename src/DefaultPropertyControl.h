//////////////////////////////////////////////////////////////////////
/// \class DefaultPropertyControl
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Provides a default implementation of the \c IPropertyControl interface</em>
///
/// TODO
///
/// \sa IPropertyControl
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IPropertyControl.h"
#include "ExplorerListView.h"


#ifdef INCLUDESUBITEMCALLBACKCODE
	class ATL_NO_VTABLE DefaultPropertyControl : 
			public CComObjectRootEx<CComSingleThreadModel>,
			public IPropertyControl
	{
		friend class ExplorerListView;

	public:
		#ifndef DOXYGEN_SHOULD_SKIP_THIS
			BEGIN_COM_MAP(DefaultPropertyControl)
				COM_INTERFACE_ENTRY_IID(IID_IPropertyControlBase, IPropertyControlBase)
				COM_INTERFACE_ENTRY_IID(IID_IPropertyControl, IPropertyControl)
			END_COM_MAP()

			DECLARE_PROTECT_FINAL_CONSTRUCT()
		#endif

		/// \brief <em>Sets the owner of this object</em>
		///
		/// \param[in] pOwner The owner to set.
		///
		/// \sa Properties::pOwnerExLvw
		void SetOwner(__in_opt ExplorerListView* pOwner);

	protected:
		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IPropertyControlBase
		///
		//@{
		virtual HRESULT STDMETHODCALLTYPE Initialize(LPUNKNOWN, PROPDESC_CONTROL_TYPE);
		virtual HRESULT STDMETHODCALLTYPE GetSize(PROPCTL_RECT_TYPE /*unknown1*/, HDC /*hDC*/, SIZE const* /*pUnknown2*/, LPSIZE /*pUnknown3*/);
		virtual HRESULT STDMETHODCALLTYPE SetWindowTheme(LPCWSTR, LPCWSTR);
		virtual HRESULT STDMETHODCALLTYPE SetFont(HFONT hFont);
		virtual HRESULT STDMETHODCALLTYPE SetTextColor(COLORREF /*color*/);
		virtual HRESULT STDMETHODCALLTYPE GetFlags(PINT);
		virtual HRESULT STDMETHODCALLTYPE SetFlags(int /*flags*/, int /*mask*/);
		virtual HRESULT STDMETHODCALLTYPE AdjustWindowRectPCB(HWND /*hWndListView*/, LPRECT /*pItemRectangle*/, RECT const* /*pUnknown1*/, int /*unknown2*/);
		virtual HRESULT STDMETHODCALLTYPE SetValue(LPUNKNOWN);
		virtual HRESULT STDMETHODCALLTYPE InvokeDefaultAction(void);
		virtual HRESULT STDMETHODCALLTYPE Destroy(void);
		virtual HRESULT STDMETHODCALLTYPE SetFormatFlags(int);
		virtual HRESULT STDMETHODCALLTYPE GetFormatFlags(PINT);
		//@}
		//////////////////////////////////////////////////////////////////////

		//////////////////////////////////////////////////////////////////////
		/// \name Implementation of IPropertyControl
		///
		//@{
		virtual HRESULT STDMETHODCALLTYPE GetValue(REFIID /*requiredInterface*/, LPVOID* /*ppObject*/);
		virtual HRESULT STDMETHODCALLTYPE Create(HWND hWndParent, RECT const* pItemRectangle, RECT const* /*pUnknown1*/, int unknown2);
		virtual HRESULT STDMETHODCALLTYPE SetPosition(RECT const*, RECT const*);
		virtual HRESULT STDMETHODCALLTYPE IsModified(LPBOOL /*pModified*/);
		virtual HRESULT STDMETHODCALLTYPE SetModified(BOOL /*modified*/);
		virtual HRESULT STDMETHODCALLTYPE ValidationFailed(LPCWSTR);
		virtual HRESULT STDMETHODCALLTYPE GetState(PINT /*pState*/);
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds the object's properties' settings</em>
		struct Properties
		{
			/// \brief <em>The \c ExplorerListView object that owns this object</em>
			///
			/// \sa SetOwner
			ExplorerListView* pOwnerExLvw;
			/// \brief <em>The edit control used for sub-item label-editing</em>
			CEdit editControl;

			Properties()
			{
				pOwnerExLvw = NULL;
			}

			~Properties();
		} properties;
	};     // DefaultPropertyControl
#endif