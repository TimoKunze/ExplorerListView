//////////////////////////////////////////////////////////////////////
/// \class DefaultDrawPropertyControl
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Provides a default implementation of the \c IDrawPropertyControl interface</em>
///
/// TODO
///
/// \sa IDrawPropertyControl
//////////////////////////////////////////////////////////////////////


#pragma once

#include "IDrawPropertyControl.h"
#include "ExplorerListView.h"


#ifdef INCLUDESUBITEMCALLBACKCODE
	class ATL_NO_VTABLE DefaultDrawPropertyControl : 
			public CComObjectRootEx<CComSingleThreadModel>,
			public IDrawPropertyControl
	{
		friend class ExplorerListView;

	public:
		#ifndef DOXYGEN_SHOULD_SKIP_THIS
			BEGIN_COM_MAP(DefaultDrawPropertyControl)
				COM_INTERFACE_ENTRY_IID(IID_IPropertyControlBase, IPropertyControlBase)
				COM_INTERFACE_ENTRY_IID(IID_IDrawPropertyControl, IDrawPropertyControl)
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
		virtual HRESULT STDMETHODCALLTYPE SetFont(HFONT /*hFont*/);
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
		/// \name Implementation of IDrawPropertyControl
		///
		//@{
		virtual HRESULT STDMETHODCALLTYPE GetDrawFlags(PINT);
		virtual HRESULT STDMETHODCALLTYPE WindowlessDraw(HDC /*hDC*/, RECT const* /*pDrawingRectangle*/, int /*unknown*/);
		virtual HRESULT STDMETHODCALLTYPE HasVisibleContent(void);
		virtual HRESULT STDMETHODCALLTYPE GetDisplayText(LPWSTR*);
		virtual HRESULT STDMETHODCALLTYPE GetTooltipInfo(HDC, SIZE const*, PINT);
		//@}
		//////////////////////////////////////////////////////////////////////

		/// \brief <em>Holds the object's properties' settings</em>
		struct Properties
		{
			/// \brief <em>The \c ExplorerListView object that owns this object</em>
			///
			/// \sa SetOwner
			ExplorerListView* pOwnerExLvw;

			Properties()
			{
				pOwnerExLvw = NULL;
			}

			~Properties();
		} properties;
	};     // DefaultDrawPropertyControl
#endif