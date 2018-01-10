// ExplorerListView.cpp: Wraps and extends SysListView32.

#include "stdafx.h"
#include "ExplorerListView.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
IMEModeConstants ExplorerListView::IMEFlags::chineseIMETable[10] = {
    imeOff,
    imeOff,
    imeOff,
    imeOff,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOn,
    imeOff
};

IMEModeConstants ExplorerListView::IMEFlags::japaneseIMETable[10] = {
    imeDisable,
    imeDisable,
    imeOff,
    imeOff,
    imeHiragana,
    imeHiragana,
    imeKatakana,
    imeKatakanaHalf,
    imeAlphaFull,
    imeAlpha
};

IMEModeConstants ExplorerListView::IMEFlags::koreanIMETable[10] = {
    imeDisable,
    imeDisable,
    imeAlpha,
    imeAlpha,
    imeHangulFull,
    imeHangul,
    imeHangulFull,
    imeHangul,
    imeAlphaFull,
    imeAlpha
};

FONTDESC ExplorerListView::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
ExplorerListView::ExplorerListView() :
    containedEdit(WC_EDIT, this, 1),
    containedSysHeader32(WC_HEADER, this, 2)
{
	properties.font.InitializePropertyWatcher(this, DISPID_EXLVW_FONT);
	properties.hotMouseIcon.InitializePropertyWatcher(this, DISPID_EXLVW_HOTMOUSEICON);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_EXLVW_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	columnUnderMouse = -1;
	itemUnderMouse.iItem = -1;
	itemUnderMouse.iGroup = 0;
	subItemUnderMouse = -1;
	hEditBackColorBrush = NULL;
	hBuiltInStateImageList = NULL;
	hBuiltInHeaderStateImageList = NULL;
	nextGroupID = 0;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
}
#pragma warning(pop)


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ExplorerListView::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IExplorerListView, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ExplorerListView::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<ExplorerListView>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("EmptyMarkupText"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_EmptyMarkupText(bstr);
		VariantClear(&propertyValue);

		bstr.Empty();
		hr = pPropertyBag->Read(OLESTR("FooterIntroText"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_FooterIntroText(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP ExplorerListView::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<ExplorerListView>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		CComBSTR bstr = L"";
		if(properties.pEmptyMarkupText) {
			bstr = properties.pEmptyMarkupText;
		}
		bstr.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("EmptyMarkupText"), &propertyValue);
		VariantClear(&propertyValue);

		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.footerIntroText.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("FooterIntroText"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP ExplorerListView::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 59 VT_I4 properties...
	pSize->QuadPart += 59 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 40 VT_BOOL properties...
	pSize->QuadPart += 40 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 2 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + CComBSTR(properties.pEmptyMarkupText).ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.footerIntroText.ByteLength() + sizeof(OLECHAR);

	// ...and 3 VT_DISPATCH properties
	pSize->QuadPart += 3 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.hotMouseIcon.pPictureDisp) {
		if(FAILED(properties.hotMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.hotMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP ExplorerListView::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	// NOTE: We neither support legacy streams nor streams with a version < 0x0102.

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x03030303/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version < 0x0102 || version > 0x0104) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version > 0x0102) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL absoluteBkImagePosition = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowHeaderDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowLabelEditing = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysShowSelection = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AutoArrangeItemsConstants autoArrangeItems = static_cast<AutoArrangeItemsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG bkImagePositionX = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG bkImagePositionY = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BkImageStyleConstants bkImageStyle = static_cast<BkImageStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL blendSelectionLasso = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL borderSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	CallBackMaskConstants callBackMask = static_cast<CallBackMaskConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL clickableColumnHeaders = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragScrollTimeBase = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR editBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR editForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG editHoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants editIMEMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR emptyMarkupText;
	if(FAILED(hr = emptyMarkupText.ReadFromStream(pStream))) {
		return hr;
	}
	if(!emptyMarkupText) {
		emptyMarkupText = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG filterChangedTimeout = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	FullRowSelectConstants fullRowSelect = frsExtendedMode;
	if(version >= 0x0104) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		fullRowSelect = static_cast<FullRowSelectConstants>(propertyValue.lVal);
	} else {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		fullRowSelect = (VARIANTBOOL2BOOL(propertyValue.boolVal) ? frsNormalMode : frsDisabled);
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL gridLines = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR groupFooterForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR groupHeaderForeColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS groupMarginBottom = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS groupMarginLeft = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS groupMarginRight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS groupMarginTop = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	SortOrderConstants groupSortOrder = static_cast<SortOrderConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL headerFullDragging = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL headerHotTracking = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG headerHoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL hideLabels = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR hotForeColor = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pHotMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pHotMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants hotMousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL hotTracking = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hotTrackingHoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	IMEModeConstants IMEMode = static_cast<IMEModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR insertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemActivationModeConstants itemActivationMode = static_cast<ItemActivationModeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemAlignmentConstants itemAlignment = static_cast<ItemAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ItemBoundingBoxDefinitionConstants itemBoundingBoxDefinition = static_cast<ItemBoundingBoxDefinitionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS itemHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL labelWrap = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR outlineColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL ownerDrawn = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL regional = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ScrollBarsConstants scrollBars = static_cast<ScrollBarsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showFilterBar = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showGroups = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showStateImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showSubItemImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL simpleSelect = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL singleRow = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL snapToGrid = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	SortOrderConstants sortOrder = static_cast<SortOrderConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR textBackColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG tileViewItemLines = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS tileViewLabelMarginBottom = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS tileViewLabelMarginLeft = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS tileViewLabelMarginRight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS tileViewLabelMarginTop = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ToolTipsConstants toolTips = static_cast<ToolTipsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	UnderlinedItemsConstants underlinedItems = static_cast<UnderlinedItemsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useWorkAreas = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ViewConstants view = static_cast<ViewConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG virtualItemCount = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL virtualMode = propertyValue.boolVal;

	VARIANT_BOOL autoSizeColumns = VARIANT_FALSE;
	BackgroundDrawModeConstants backgroundDrawMode = bdmNormal;
	VARIANT_BOOL checkItemOnSelect = VARIANT_FALSE;
	ColumnHeaderVisibilityConstants columnHeaderVisibility = chvVisibleInDetailsView;
	VARIANT_BOOL drawImagesAsynchronously = VARIANT_FALSE;
	AlignmentConstants emptyMarkupTextAlignment = alCenter;
	CComBSTR footerIntroText;
	OLEDragImageStyleConstants headerOLEDragImageStyle = odistClassic;
	VARIANT_BOOL includeHeaderInTabOrder = VARIANT_FALSE;
	VARIANT_BOOL justifyIconColumns = VARIANT_FALSE;
	LONG minItemRowsVisibleInGroups = 0;
	OLEDragImageStyleConstants oleDragImageStyle = odistClassic;
	VARIANT_BOOL resizableColumns = VARIANT_TRUE;
	OLE_COLOR selectedColumnBackColor = static_cast<OLE_COLOR>(-1);
	VARIANT_BOOL showHeaderChevron = VARIANT_FALSE;
	VARIANT_BOOL showHeaderStateImages = VARIANT_FALSE;
	OLE_COLOR tileViewSubItemForeColor = static_cast<OLE_COLOR>(-1);
	OLE_YSIZE_PIXELS tileViewTileHeight = -1;
	OLE_XSIZE_PIXELS tileViewTileWidth = -1;
	VARIANT_BOOL useMinColumnWidths = VARIANT_FALSE;
	if(version >= 0x0102) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		autoSizeColumns = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		backgroundDrawMode = static_cast<BackgroundDrawModeConstants>(propertyValue.lVal);
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		checkItemOnSelect = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		columnHeaderVisibility = static_cast<ColumnHeaderVisibilityConstants>(propertyValue.lVal);
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		drawImagesAsynchronously = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		emptyMarkupTextAlignment = static_cast<AlignmentConstants>(propertyValue.lVal);
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		justifyIconColumns = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		resizableColumns = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		showHeaderChevron = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		showHeaderStateImages = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		tileViewTileHeight = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		tileViewTileWidth = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		useMinColumnWidths = propertyValue.boolVal;
		if(version >= 0x0103) {
			if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
				return hr;
			}
			if(FAILED(hr = footerIntroText.ReadFromStream(pStream))) {
				return hr;
			}
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			headerOLEDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
				return hr;
			}
			includeHeaderInTabOrder = propertyValue.boolVal;
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			minItemRowsVisibleInGroups = propertyValue.lVal;
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			selectedColumnBackColor = propertyValue.lVal;
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
				return hr;
			}
			tileViewSubItemForeColor = propertyValue.lVal;
		}
	}


	hr = put_AbsoluteBkImagePosition(absoluteBkImagePosition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowHeaderDragDrop(allowHeaderDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AllowLabelEditing(allowLabelEditing);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysShowSelection(alwaysShowSelection);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoArrangeItems(autoArrangeItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoSizeColumns(autoSizeColumns);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackgroundDrawMode(backgroundDrawMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BkImagePositionX(bkImagePositionX);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BkImagePositionY(bkImagePositionY);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BkImageStyle(bkImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BlendSelectionLasso(blendSelectionLasso);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderSelect(borderSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CallBackMask(callBackMask);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CheckItemOnSelect(checkItemOnSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ClickableColumnHeaders(clickableColumnHeaders);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ColumnHeaderVisibility(columnHeaderVisibility);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragScrollTimeBase(dragScrollTimeBase);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DrawImagesAsynchronously(drawImagesAsynchronously);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditBackColor(editBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditForeColor(editForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditHoverTime(editHoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EditIMEMode(editIMEMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EmptyMarkupText(emptyMarkupText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_EmptyMarkupTextAlignment(emptyMarkupTextAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FilterChangedTimeout(filterChangedTimeout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FooterIntroText(footerIntroText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FullRowSelect(fullRowSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GridLines(gridLines);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupFooterForeColor(groupFooterForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupHeaderForeColor(groupHeaderForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupMarginBottom(groupMarginBottom);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupMarginLeft(groupMarginLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupMarginRight(groupMarginRight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupMarginTop(groupMarginTop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_GroupSortOrder(groupSortOrder);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HeaderFullDragging(headerFullDragging);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HeaderHotTracking(headerHotTracking);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HeaderHoverTime(headerHoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HeaderOLEDragImageStyle(headerOLEDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HideLabels(hideLabels);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotForeColor(hotForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotMouseIcon(pHotMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotMousePointer(hotMousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotTracking(hotTracking);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotTrackingHoverTime(hotTrackingHoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IMEMode(IMEMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_IncludeHeaderInTabOrder(includeHeaderInTabOrder);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkColor(insertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemActivationMode(itemActivationMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemAlignment(itemAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemBoundingBoxDefinition(itemBoundingBoxDefinition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ItemHeight(itemHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_JustifyIconColumns(justifyIconColumns);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_LabelWrap(labelWrap);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinItemRowsVisibleInGroups(minItemRowsVisibleInGroups);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MultiSelect(multiSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OutlineColor(outlineColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OwnerDrawn(ownerDrawn);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Regional(regional);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ResizableColumns(resizableColumns);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollBars(scrollBars);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SelectedColumnBackColor(selectedColumnBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowFilterBar(showFilterBar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowGroups(showGroups);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowHeaderChevron(showHeaderChevron);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowHeaderStateImages(showHeaderStateImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowStateImages(showStateImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowSubItemImages(showSubItemImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SimpleSelect(simpleSelect);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SingleRow(singleRow);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SnapToGrid(snapToGrid);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SortOrder(sortOrder);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextBackColor(textBackColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewItemLines(tileViewItemLines);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewLabelMarginBottom(tileViewLabelMarginBottom);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewLabelMarginLeft(tileViewLabelMarginLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewLabelMarginRight(tileViewLabelMarginRight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewLabelMarginTop(tileViewLabelMarginTop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewSubItemForeColor(tileViewSubItemForeColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewTileHeight(tileViewTileHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TileViewTileWidth(tileViewTileWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ToolTips(toolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UnderlinedItems(underlinedItems);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseMinColumnWidths(useMinColumnWidths);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseWorkAreas(useWorkAreas);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_View(view);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VirtualItemCount(VARIANT_FALSE, VARIANT_TRUE, virtualItemCount);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VirtualMode(virtualMode);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ExplorerListView::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x03030303/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0104;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.absoluteBkImagePosition);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowHeaderDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowLabelEditing);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.autoArrangeItems;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.bkImagePositionX;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.bkImagePositionY;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.bkImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.blendSelectionLasso);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.borderSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.callBackMask;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.clickableColumnHeaders);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragScrollTimeBase;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editHoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.editIMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	VARTYPE vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = CComBSTR(properties.pEmptyMarkupText).WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.filterChangedTimeout;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.fullRowSelect;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.gridLines);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.groupFooterForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.groupHeaderForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.groupMarginBottom;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.groupMarginLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.groupMarginRight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.groupMarginTop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.groupSortOrder;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.headerFullDragging);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.headerHotTracking);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.headerHoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.hideLabels);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hotForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.hotMouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.hotMouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.hotMousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.hotTracking);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hotTrackingHoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.IMEMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemActivationMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemBoundingBoxDefinition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.itemHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.labelWrap);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.outlineColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.ownerDrawn);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.regional);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.scrollBars;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showFilterBar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showGroups);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showStateImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showSubItemImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.simpleSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.singleRow);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.snapToGrid);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.sortOrder;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.textBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewItemLines;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewLabelMarginBottom;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewLabelMarginLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewLabelMarginRight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewLabelMarginTop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.toolTips;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.underlinedItems;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useWorkAreas);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.view;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.virtualItemCount;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.virtualMode);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoSizeColumns);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backgroundDrawMode;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.checkItemOnSelect);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.columnHeaderVisibility;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.drawImagesAsynchronously);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.emptyMarkupTextAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.justifyIconColumns);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.resizableColumns);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showHeaderChevron);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showHeaderStateImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.tileViewTileHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewTileWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useMinColumnWidths);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0103 starts here
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.footerIntroText.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.headerOLEDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.includeHeaderInTabOrder);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.minItemRowsVisibleInGroups;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.selectedColumnBackColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.tileViewSubItemForeColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ExplorerListView::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ExplorerListView>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ExplorerListView::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ExplorerListView ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ExplorerListView::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ExplorerListView::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ExplorerListView>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

STDMETHODIMP ExplorerListView::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL ExplorerListView::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		//ATLASSERT(dialogCode == (pMessage->hwnd == containedEdit.m_hWnd ? (DLGC_WANTALLKEYS | DLGC_HASSETSEL) : (DLGC_WANTARROWS | (pMessage->hwnd == containedSysHeader32.m_hWnd ? DLGC_WANTTAB : DLGC_WANTCHARS))));
		/*if(dialogCode & DLGC_WANTTAB) {
			if(pMessage->wParam == VK_TAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}*/
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(((dialogCode & DLGC_WANTARROWS) == DLGC_WANTARROWS) || ((dialogCode & DLGC_WANTALLKEYS) == DLGC_WANTALLKEYS)) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
			case VK_TAB:
				if((pMessage->message == WM_KEYDOWN) && properties.includeHeaderInTabOrder && IsComctl32Version610OrNewer() && containedSysHeader32.IsWindowVisible()) {
					// insert the header control into the tab order
					HWND hWndFocus = GetFocus();
					if(hWndFocus == *this) {
						if(!(GetKeyState(VK_SHIFT) & 0x8000)) {
							// not pressed
							containedSysHeader32.SetFocus();
							hReturnValue = S_FALSE;
							return TRUE;
						}
					} else if(hWndFocus == containedSysHeader32.m_hWnd) {
						if(GetKeyState(VK_SHIFT) & 0x8000) {
							// pressed
							this->SetFocus();
							hReturnValue = S_FALSE;
							return TRUE;
						}
					}
				}
				break;
		}
	}
	return CComControl<ExplorerListView>::PreTranslateAccelerator(pMessage, hReturnValue);
}

HIMAGELIST ExplorerListView::CreateLegacyDragImage(LVITEMINDEX& itemIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	/********************************************************************************************************
	 * Known problems:                                                                                      *
	 * - We assume that the selection background goes over icon and label for all themes.                   *
	 * - We use hardcoded margins.                                                                          *
	 * - Extended Tile View is not fully supported (displaying the details in two columns doesn't work).    *
	 ********************************************************************************************************/

	TCHAR pItemTextBuffer[MAX_ITEMTEXTLENGTH + 1];
	TCHAR pSubItemTextBuffer[MAX_ITEMTEXTLENGTH + 1];

	// retrieve item details
	LVITEM item = {0};
	item.iItem = itemIndex.iItem;
	item.iGroup = itemIndex.iGroup;
	item.cchTextMax = MAX_ITEMTEXTLENGTH;
	item.pszText = pItemTextBuffer;
	item.mask = LVIF_IMAGE | LVIF_STATE | LVIF_TEXT;
	item.stateMask = LVIS_FOCUSED | LVIS_SELECTED | LVIS_OVERLAYMASK;
	SendMessage(LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item));
	if(!item.pszText) {
		item.pszText = TEXT("");
	}
	BOOL itemIsSelected = ((item.state & LVIS_SELECTED) == LVIS_SELECTED);
	BOOL itemIsFocused = ((item.state & LVIS_FOCUSED) == LVIS_FOCUSED);
	int iconToDraw = item.iImage;
	CComPtr<IListViewItem> pItem = NULL;
	VARIANT vTileViewColumns;
	VariantInit(&vTileViewColumns);
	int lines = 0;
	int maxDetailLinesBySettings = 0;
	int numberOfMainTextLines = 0;
	int lineHeight = 0;
	LVTILEVIEWINFO tileViewInfo = {0};
	int columnCount = 0;
	LPRECT pSubItemImageRectangles = NULL;
	LPBOOL pSubItemImageHasAlpha = NULL;

	// retrieve window details
	BOOL hasFocus = (GetFocus() == *this);
	DWORD style = GetExStyle();
	DWORD textDrawStyle = DT_EDITCONTROL | DT_NOPREFIX;
	if(style & WS_EX_RTLREADING) {
		textDrawStyle |= DT_RTLREADING;
	}
	BOOL layoutRTL = ((style & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	DWORD extendedStyle = static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0));
	BOOL fullRowSelect = ((extendedStyle & LVS_EX_FULLROWSELECT) == LVS_EX_FULLROWSELECT);
	ViewConstants currentView = vIcons;
	get_View(&currentView);
	HIMAGELIST hSourceImageList = NULL;
	switch(currentView) {
		case vDetails:
			hSourceImageList = cachedSettings.hSmallImageList;
			textDrawStyle |= DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
			break;
		case vIcons:
			hSourceImageList = cachedSettings.hLargeImageList;
			textDrawStyle |= DT_CENTER | DT_WORDBREAK | DT_END_ELLIPSIS;
			break;
		case vList:
			hSourceImageList = cachedSettings.hSmallImageList;
			textDrawStyle |= DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
			break;
		case vSmallIcons:
			hSourceImageList = cachedSettings.hSmallImageList;
			textDrawStyle |= DT_SINGLELINE;
			break;
		case vTiles:
			hSourceImageList = cachedSettings.hExtraLargeImageList;
			textDrawStyle |= DT_END_ELLIPSIS;
			break;
		case vExtendedTiles:
			hSourceImageList = cachedSettings.hExtraLargeImageList;
			textDrawStyle |= DT_END_ELLIPSIS;
			break;
	}
	SIZE imageSize = {0};
	BOOL subItemImages = FALSE;
	if(hSourceImageList) {
		if(currentView == vDetails) {
			subItemImages = ((extendedStyle & LVS_EX_SUBITEMIMAGES) == LVS_EX_SUBITEMIMAGES);
		}

		ImageList_GetIconSize(hSourceImageList, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
	}

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	CFontHandle font = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
	HFONT hPreviousFont = NULL;
	if(!font.IsNull()) {
		hPreviousFont = memoryDC.SelectFont(font);
	}

	// prepare themed item drawing
	CTheme themingEngine;
	BOOL themedListItems = FALSE;
	int themeState = LISS_NORMAL;
	if(itemIsSelected) {
		if(hasFocus) {
			themeState = LISS_SELECTED;
		} else {
			themeState = LISS_SELECTEDNOTFOCUS;
		}
	}
	if(flags.usingThemes) {
		themingEngine.OpenThemeData(*this, VSCLASS_LISTVIEW);
		/* We use LISS_SELECTED here, because it's more likely this one is defined and we don't want a mixture
		   of themed and non-themed items in drag images. What we're doing with LISS_NORMAL should work
		   regardless whether LISS_NORMAL is defined. */
		themedListItems = themingEngine.IsThemePartDefined(LVP_LISTITEM, LISS_SELECTED/*themeState*/);
		if(themedListItems) {
			if(currentView == vSmallIcons) {
				textDrawStyle |= DT_VCENTER;
			}
		}
	}

	// calculate the bounding rectangles of the various item parts
	WTL::CRect itemBoundingRect;
	WTL::CRect selectionBoundingRect;
	WTL::CRect focusRect;
	WTL::CRect labelBoundingRect;
	WTL::CRect iconAreaBoundingRect;
	WTL::CRect iconBoundingRect;

	if(IsComctl32Version610OrNewer() && SendMessage(LVM_ISGROUPVIEWENABLED, 0, 0) && (GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		labelBoundingRect.left = LVIR_LABEL;
		SendMessage(LVM_GETITEMINDEXRECT, reinterpret_cast<WPARAM>(&itemIndex), reinterpret_cast<LPARAM>(&labelBoundingRect));
		selectionBoundingRect.left = LVIR_SELECTBOUNDS;
		SendMessage(LVM_GETITEMINDEXRECT, reinterpret_cast<WPARAM>(&itemIndex), reinterpret_cast<LPARAM>(&selectionBoundingRect));
	} else {
		labelBoundingRect.left = LVIR_LABEL;
		SendMessage(LVM_GETITEMRECT, itemIndex.iItem, reinterpret_cast<LPARAM>(&labelBoundingRect));
		selectionBoundingRect.left = LVIR_SELECTBOUNDS;
		SendMessage(LVM_GETITEMRECT, itemIndex.iItem, reinterpret_cast<LPARAM>(&selectionBoundingRect));
	}
	itemBoundingRect = selectionBoundingRect;
	if(!themedListItems) {
		if(currentView == vList) {
			// the width of labelBoundingRect may be wrong
			if(lstrlen(item.pszText) > 0) {
				WTL::CRect rc = labelBoundingRect;
				memoryDC.DrawText(item.pszText, lstrlen(item.pszText), &rc, textDrawStyle | DT_CALCRECT);
				// TODO: Don't use hard-coded margins
				labelBoundingRect.right = labelBoundingRect.left + rc.Width() + 4;
			}
		}
	}
	if(currentView == vTiles || currentView == vExtendedTiles) {
		// the height of labelBoundingRect may be wrong

		// calculate the line height
		WTL::CRect rc(0, 0, 30, 200000);
		LPTSTR pDummy = TEXT("A\r\nB");
		// TODO: DrawThemeText() doesn't set the rectangle's height
		/*if(themedListItems) {
			CT2W converter(pDummy);
			LPWSTR pLabelText = converter;
			themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_WORDBREAK | DT_CALCRECT, 0, &rc);
		} else {*/
			memoryDC.DrawText(pDummy, lstrlen(pDummy), &rc, textDrawStyle | DT_WORDBREAK | DT_CALCRECT);
		//}
		lineHeight = rc.Height() / 2;

		// retrieve the maximum number of lines
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_COLUMNS | LVTVIM_TILESIZE;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		// retrieve the number of sub-items to display
		pItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
		ATLASSUME(pItem);
		ATLVERIFY(SUCCEEDED(pItem->get_TileViewColumns(&vTileViewColumns)));
		ATLASSERT(vTileViewColumns.vt & VT_ARRAY);
		maxDetailLinesBySettings = 0;
		if(vTileViewColumns.parray) {
			LONG l = 0;
			SafeArrayGetLBound(vTileViewColumns.parray, 1, &l);
			LONG u = 0;
			SafeArrayGetUBound(vTileViewColumns.parray, 1, &u);
			maxDetailLinesBySettings = min(u - l + 1, tileViewInfo.cLines);
		}

		// count the lines
		numberOfMainTextLines = 1;
		lines = 1;
		CAtlString itemText = item.pszText;
		// the item's main text may consist of up to two lines, so check this
		if(itemText.Find(TEXT("\r\n")) != -1) {
			++lines;
			++numberOfMainTextLines;
		}
		int maxLines = min(maxDetailLinesBySettings + numberOfMainTextLines, tileViewInfo.cLines + 1);
		if(currentView == vExtendedTiles) {
			int maxLinesByHeight = labelBoundingRect.Height() / lineHeight;
			maxLines = min(maxLines, maxLinesByHeight);
		}

		LVITEM subItem = {0};
		subItem.cchTextMax = MAX_ITEMTEXTLENGTH;
		for(LONG line = 0; line < maxDetailLinesBySettings; ++line) {
			subItem.pszText = pSubItemTextBuffer;
			ATLVERIFY(SUCCEEDED(SafeArrayGetElement(vTileViewColumns.parray, &line, &subItem.iSubItem)));
			SendMessage(LVM_GETITEMTEXT, itemIndex.iItem, reinterpret_cast<LPARAM>(&subItem));
			if(subItem.pszText) {
				if(lstrlen(subItem.pszText) > 0) {
					++lines;
				}
			}
		}
		lines = min(lines, maxLines);

		// calculate the real height of labelBoundingRect
		if(currentView == vExtendedTiles && !themedListItems) {
			labelBoundingRect.bottom = selectionBoundingRect.bottom;
		} else {
			labelBoundingRect.bottom = labelBoundingRect.top + lines * lineHeight;
		}

		if(!themedListItems && currentView != vExtendedTiles) {
			if(!(tileViewInfo.dwFlags & LVTVIF_FIXEDWIDTH) && !fullRowSelect) {
				// the width may be wrong, too
				int remainingLines = lines;
				WTL::CRect textAreaRect;
				rc = labelBoundingRect;
				rc.bottom = rc.top + lineHeight;
				if((numberOfMainTextLines == 1) || (remainingLines == 1)) {
					memoryDC.DrawText(item.pszText, lstrlen(item.pszText), &rc, textDrawStyle | DT_WORDBREAK | DT_SINGLELINE | DT_CALCRECT);
					textAreaRect.UnionRect(&textAreaRect, &rc);
					--remainingLines;
				} else {
					int i = itemText.Find(TEXT("\r\n"));
					ATLASSERT(i > 0);
					CAtlString firstLine = itemText.Left(i);
					CAtlString secondLine = itemText.Mid(i + 2);
					memoryDC.DrawText(firstLine, lstrlen(firstLine), &rc, textDrawStyle | DT_WORDBREAK | DT_CALCRECT);
					textAreaRect.UnionRect(&textAreaRect, &rc);
					--remainingLines;
					rc = labelBoundingRect;
					rc.top = labelBoundingRect.top + lineHeight;
					rc.bottom = rc.top + lineHeight;
					memoryDC.DrawText(secondLine, lstrlen(secondLine), &rc, textDrawStyle | DT_WORDBREAK | DT_SINGLELINE | DT_CALCRECT);
					textAreaRect.UnionRect(&textAreaRect, &rc);
					--remainingLines;
				}
				int numberOfDetailLines = min(remainingLines, maxDetailLinesBySettings);
				int nonEmptyLine = 0;
				for(LONG line = 0; nonEmptyLine < numberOfDetailLines; ++line) {
					subItem.pszText = pSubItemTextBuffer;
					ATLVERIFY(SUCCEEDED(SafeArrayGetElement(vTileViewColumns.parray, &line, &subItem.iSubItem)));
					SendMessage(LVM_GETITEMTEXT, itemIndex.iItem, reinterpret_cast<LPARAM>(&subItem));
					if(subItem.pszText) {
						if(lstrlen(subItem.pszText) > 0) {
							rc = labelBoundingRect;
							rc.top = labelBoundingRect.top + (lines - remainingLines) * lineHeight;
							rc.bottom = rc.top + lineHeight;

							memoryDC.DrawText(subItem.pszText, lstrlen(subItem.pszText), &rc, textDrawStyle | DT_SINGLELINE | DT_CALCRECT);
							textAreaRect.UnionRect(&textAreaRect, &rc);
							--remainingLines;
							++nonEmptyLine;
						}
					}
				}

				labelBoundingRect.right = textAreaRect.right;
			}
		}
	}
	if(currentView != vIcons && currentView != vExtendedTiles) {
		// center the label rectangle vertically
		int cy = labelBoundingRect.Height();
		labelBoundingRect.top = itemBoundingRect.top + (itemBoundingRect.Height() - cy) / 2;
		labelBoundingRect.bottom = labelBoundingRect.top + cy;
	}
	if(themedListItems) {
		if(currentView == vDetails) {
			if(!fullRowSelect) {
				// TODO: Don't use hard-coded margins
				selectionBoundingRect.right -= 2;
				labelBoundingRect.right = min(labelBoundingRect.right, selectionBoundingRect.right);
			}
		}
	} else {
		// selectionBoundingRect includes the icon, so correct it
		if(currentView == vDetails) {
			// labelBoundingRect covers the entire first column and nothing more
			selectionBoundingRect.left = labelBoundingRect.left;
			if(!fullRowSelect) {
				labelBoundingRect.right = selectionBoundingRect.right;
			}
		} else {
			selectionBoundingRect = labelBoundingRect;
		}
	}
	if(hSourceImageList) {
		iconAreaBoundingRect.left = LVIR_ICON;
		if(IsComctl32Version610OrNewer() && SendMessage(LVM_ISGROUPVIEWENABLED, 0, 0) && (GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
			SendMessage(LVM_GETITEMINDEXRECT, reinterpret_cast<WPARAM>(&itemIndex), reinterpret_cast<LPARAM>(&iconAreaBoundingRect));
		} else {
			SendMessage(LVM_GETITEMRECT, itemIndex.iItem, reinterpret_cast<LPARAM>(&iconAreaBoundingRect));
		}
		if(themedListItems) {
			if(currentView == vSmallIcons) {
				// TODO: Don't use hard-coded margins
				iconAreaBoundingRect.left += 4;
			}
		}
		if(currentView == vTiles || currentView == vExtendedTiles) {
			if(iconAreaBoundingRect.Height() < labelBoundingRect.Height()) {
				iconAreaBoundingRect.top = labelBoundingRect.top;
				iconAreaBoundingRect.bottom = labelBoundingRect.bottom;
			}
			// TODO: Don't use hard-coded margins
			iconAreaBoundingRect.right = labelBoundingRect.left - 4;
			iconAreaBoundingRect.left = iconAreaBoundingRect.right - imageSize.cx;
		}
		if(currentView != vIcons) {
			// center the icon area rectangle vertically
			int cy = iconAreaBoundingRect.Height();
			iconAreaBoundingRect.top = itemBoundingRect.top + (itemBoundingRect.Height() - cy) / 2;
			iconAreaBoundingRect.bottom = iconAreaBoundingRect.top + cy;
		}
		iconBoundingRect = iconAreaBoundingRect;
		if(currentView == vIcons) {
			// center the icon horizontally
			// TODO: Don't use hard-coded margins
			iconBoundingRect.top += 2;
			iconBoundingRect.left += (iconAreaBoundingRect.Width() - imageSize.cx) / 2;
			iconBoundingRect.bottom = iconBoundingRect.top + imageSize.cy;
			iconBoundingRect.right = iconBoundingRect.left + imageSize.cx;
		} else {
			// center the icon vertically
			iconBoundingRect.top += (iconAreaBoundingRect.Height() - imageSize.cy) / 2;
			iconBoundingRect.bottom = iconBoundingRect.top + imageSize.cy;
			iconBoundingRect.right = iconBoundingRect.left + imageSize.cx;
		}
	}
	itemBoundingRect.SetRectEmpty();
	itemBoundingRect.left = LVIR_BOUNDS;
	if(IsComctl32Version610OrNewer() && SendMessage(LVM_ISGROUPVIEWENABLED, 0, 0) && (GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
		SendMessage(LVM_GETITEMINDEXRECT, reinterpret_cast<WPARAM>(&itemIndex), reinterpret_cast<LPARAM>(&itemBoundingRect));
	} else {
		SendMessage(LVM_GETITEMRECT, itemIndex.iItem, reinterpret_cast<LPARAM>(&itemBoundingRect));
	}
	if(hSourceImageList) {
		itemBoundingRect.UnionRect(&itemBoundingRect, &iconAreaBoundingRect);
	}
	itemBoundingRect.UnionRect(&itemBoundingRect, &labelBoundingRect);
	focusRect = selectionBoundingRect;
	if(pBoundingRectangle) {
		*pBoundingRectangle = itemBoundingRect;
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = itemBoundingRect.left;
		pUpperLeftPoint->y = itemBoundingRect.top;
	}
	dragImageSize.cx = itemBoundingRect.Width();
	dragImageSize.cy = itemBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = itemBoundingRect.left;
	offset.cy = itemBoundingRect.top;
	labelBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	iconAreaBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	iconBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	selectionBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	itemBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	if(itemIsSelected) {
		memoryDC.SetBkColor(GetSysColor((hasFocus ? COLOR_HIGHLIGHT : COLOR_BTNFACE)));
		memoryDC.SetTextColor(GetSysColor((hasFocus ? COLOR_HIGHLIGHTTEXT : COLOR_BTNTEXT)));
	} else {
		memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
		memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
		memoryDC.SetBkMode(TRANSPARENT);
	}

	// create drag image bitmap
	/* NOTE: We prefer creating 32bpp drag images, because this improves performance of
	         ListViewItemContainer::CreateDragImage(). */
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	if(themedListItems) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&itemBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&itemBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

	// draw the selection area's background
	if(itemIsSelected) {
		if(themedListItems) {
			themingEngine.DrawThemeBackground(memoryDC, LVP_LISTITEM, themeState, &selectionBoundingRect, NULL);
		} else {
			memoryDC.FillRect(&selectionBoundingRect, (hasFocus ? COLOR_HIGHLIGHT : COLOR_BTNFACE));
		}
		maskMemoryDC.FillRect(&selectionBoundingRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
	}

	// draw the icon
	if(hSourceImageList) {
		ImageList_DrawEx(hSourceImageList, iconToDraw, memoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_NONE, CLR_NONE, ILD_NORMAL | (item.state & LVIS_OVERLAYMASK));
		if(itemIsSelected && themedListItems) {
			maskMemoryDC.FillRect(&iconBoundingRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
		} else {
			ImageList_DrawEx(hSourceImageList, iconToDraw, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_DEFAULT, CLR_DEFAULT, ILD_MASK | (item.state & LVIS_OVERLAYMASK));
			//ImageList_Draw(hSourceImageList, iconToDraw, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, ILD_MASK | (item.state & LVIS_OVERLAYMASK));
		}
	}

	// draw the text
	WTL::CRect rc = labelBoundingRect;
	if(currentView == vTiles || currentView == vExtendedTiles) {
		int remainingLines = lines;
		rc.bottom = rc.top + lineHeight;
		if(numberOfMainTextLines == 1 || remainingLines == 1) {
			if(themedListItems) {
				CT2W converter(item.pszText);
				LPWSTR pLabelText = converter;
				themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_WORDBREAK | DT_SINGLELINE, 0, &rc);
			} else {
				memoryDC.DrawText(item.pszText, lstrlen(item.pszText), &rc, textDrawStyle | DT_WORDBREAK | DT_SINGLELINE);
			}
			--remainingLines;
		} else {
			CAtlString itemText = item.pszText;
			int i = itemText.Find(TEXT("\r\n"));
			ATLASSERT(i > 0);
			CAtlString firstLine = itemText.Left(i);
			CAtlString secondLine = itemText.Mid(i + 2);
			if(themedListItems) {
				CT2W converter(firstLine);
				LPWSTR pLabelText = converter;
				themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_WORDBREAK, 0, &rc);
			} else {
				memoryDC.DrawText(firstLine, lstrlen(firstLine), &rc, textDrawStyle | DT_WORDBREAK);
			}
			--remainingLines;
			if(!itemIsSelected) {
				COLORREF bkColor = memoryDC.GetBkColor();
				for(int y = rc.top; y <= rc.bottom; ++y) {
					for(int x = rc.left; x <= rc.right; ++x) {
						if(memoryDC.GetPixel(x, y) != bkColor) {
							maskMemoryDC.SetPixelV(x, y, 0x00000000);
						}
					}
				}
			}
			rc = labelBoundingRect;
			rc.MoveToY(labelBoundingRect.top + lineHeight);
			if(themedListItems) {
				CT2W converter(secondLine);
				LPWSTR pLabelText = converter;
				themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_WORDBREAK | DT_SINGLELINE, 0, &rc);
			} else {
				memoryDC.DrawText(secondLine, lstrlen(secondLine), &rc, textDrawStyle | DT_WORDBREAK | DT_SINGLELINE);
			}
			--remainingLines;
		}
		if(!itemIsSelected) {
			COLORREF bkColor = memoryDC.GetBkColor();
			for(int y = rc.top; y <= rc.bottom; ++y) {
				for(int x = rc.left; x <= rc.right; ++x) {
					if(memoryDC.GetPixel(x, y) != bkColor) {
						maskMemoryDC.SetPixelV(x, y, 0x00000000);
					}
				}
			}
		}
		LVITEM subItem = {0};
		subItem.cchTextMax = MAX_ITEMTEXTLENGTH;
		int numberOfDetailLines = min(remainingLines, maxDetailLinesBySettings);
		int nonEmptyLine = 0;
		for(LONG line = 0; nonEmptyLine < numberOfDetailLines; ++line) {
			subItem.pszText = pSubItemTextBuffer;
			ATLVERIFY(SUCCEEDED(SafeArrayGetElement(vTileViewColumns.parray, &line, &subItem.iSubItem)));
			SendMessage(LVM_GETITEMTEXT, itemIndex.iItem, reinterpret_cast<LPARAM>(&subItem));
			if(subItem.pszText) {
				if(lstrlen(subItem.pszText) > 0) {
					rc = labelBoundingRect;
					rc.top = labelBoundingRect.top + (lines - remainingLines) * lineHeight;
					rc.bottom = rc.top + lineHeight;

					if(themedListItems) {
						CT2W converter(subItem.pszText);
						LPWSTR pLabelText = converter;
						themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle | DT_SINGLELINE, 0, &rc);
					} else {
						memoryDC.DrawText(subItem.pszText, lstrlen(subItem.pszText), &rc, textDrawStyle | DT_SINGLELINE);
					}
					--remainingLines;
					++nonEmptyLine;
					if(!itemIsSelected) {
						COLORREF bkColor = memoryDC.GetBkColor();
						for(int y = rc.top; y <= rc.bottom; ++y) {
							for(int x = rc.left; x <= rc.right; ++x) {
								if(memoryDC.GetPixel(x, y) != bkColor) {
									maskMemoryDC.SetPixelV(x, y, 0x00000000);
								}
							}
						}
					}
				}
			}
		}
	} else {
		if(currentView == vDetails) {
			if(layoutRTL) {
				// TODO: Don't use hard-coded margins
				rc.OffsetRect(1, 0);
			}
		}
		// TODO: Don't use hard-coded margins
		rc.InflateRect(-2, 0);
		if(themedListItems) {
			CT2W converter(item.pszText);
			LPWSTR pLabelText = converter;
			themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle, 0, &rc);
		} else {
			memoryDC.DrawText(item.pszText, lstrlen(item.pszText), &rc, textDrawStyle);
		}
		if(!itemIsSelected) {
			COLORREF bkColor = memoryDC.GetBkColor();
			for(int y = rc.top; y <= rc.bottom; ++y) {
				for(int x = rc.left; x <= rc.right; ++x) {
					if(memoryDC.GetPixel(x, y) != bkColor) {
						maskMemoryDC.SetPixelV(x, y, 0x00000000);
					}
				}
			}
		}

		if(currentView == vDetails && fullRowSelect) {
			columnCount = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0));
			pSubItemImageRectangles = new RECT[columnCount - 1];
			pSubItemImageHasAlpha = new BOOL[columnCount - 1];

			// draw the sub-items
			LVITEM subItem = {0};
			subItem.iItem = itemIndex.iItem;
			subItem.cchTextMax = MAX_ITEMTEXTLENGTH;
			subItem.mask = LVIF_STATE | LVIF_TEXT;
			if(subItemImages) {
				subItem.mask |= LVIF_IMAGE;
			}
			subItem.stateMask = LVIS_OVERLAYMASK;
			LVCOLUMN column = {0};
			column.mask = LVCF_FMT;

			typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
			HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
			}
			for(int subItemIndex = 1; subItemIndex < columnCount; ++subItemIndex) {
				// collect all the data we'll need
				WTL::CRect subItemRect;
				CComPtr<IListView_WIN7> pListView7 = NULL;
				CComPtr<IListView_WINVISTA> pListViewVista = NULL;
				if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
					ATLASSUME(pListView7);
					ATLVERIFY(SUCCEEDED(pListView7->GetSubItemRect(itemIndex, subItemIndex, LVIR_BOUNDS, &subItemRect)));
				} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
					ATLASSUME(pListViewVista);
					ATLVERIFY(SUCCEEDED(pListViewVista->GetSubItemRect(itemIndex, subItemIndex, LVIR_BOUNDS, &subItemRect)));
				} else {
					subItemRect.left = LVIR_BOUNDS;
					subItemRect.top = subItemIndex;
					SendMessage(LVM_GETSUBITEMRECT, itemIndex.iItem, reinterpret_cast<LPARAM>(&subItemRect));
				}
				subItemRect.OffsetRect(-offset.cx, -offset.cy);

				subItem.pszText = pSubItemTextBuffer;
				subItem.iSubItem = subItemIndex;
				SendMessage(LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&subItem));
				if(!subItem.pszText) {
					continue;
				}
				SendMessage(LVM_GETCOLUMN, subItemIndex, reinterpret_cast<LPARAM>(&column));
				switch(column.fmt & LVCFMT_JUSTIFYMASK) {
					case LVCFMT_LEFT:
						textDrawStyle &= ~(DT_CENTER | DT_RIGHT);
						textDrawStyle |= DT_LEFT;
						break;
					case LVCFMT_CENTER:
						textDrawStyle &= ~(DT_LEFT | DT_RIGHT);
						textDrawStyle |= DT_CENTER;
						break;
					case LVCFMT_RIGHT:
						textDrawStyle &= ~(DT_LEFT | DT_CENTER);
						textDrawStyle |= DT_RIGHT;
						break;
				}

				// draw the icon
				if(subItemImages) {
					iconBoundingRect = subItemRect;
					// TODO: Don't use hard-coded margins
					iconBoundingRect.right = iconBoundingRect.left + imageSize.cx + 2;

					if(RunTimeHelper::IsCommCtrl6()) {
						pSubItemImageRectangles[subItemIndex - 1] = iconBoundingRect;
						IImageList* pImgLst = NULL;
						if(pfnHIMAGELIST_QueryInterface) {
							pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
						} else {
							pImgLst = reinterpret_cast<IImageList*>(hSourceImageList);
							pImgLst->AddRef();
						}
						ATLASSUME(pImgLst);

						DWORD flags = 0;
						pImgLst->GetItemFlags(subItem.iImage, &flags);
						pImgLst->Release();
						pSubItemImageHasAlpha[subItemIndex - 1] = ((flags & ILIF_ALPHA) == ILIF_ALPHA);
					}

					// background
					if(!themedListItems) {
						memoryDC.FillRect(&iconBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
					}
					maskMemoryDC.FillRect(&iconBoundingRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

					// icon
					ImageList_DrawEx(hSourceImageList, subItem.iImage, memoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_NONE, CLR_NONE, ILD_NORMAL | (subItem.state & LVIS_OVERLAYMASK));
					if(itemIsSelected && themedListItems) {
						maskMemoryDC.FillRect(&iconBoundingRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
					} else {
						ImageList_DrawEx(hSourceImageList, subItem.iImage, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_DEFAULT, CLR_DEFAULT, ILD_MASK | (subItem.state & LVIS_OVERLAYMASK));
						//ImageList_Draw(hSourceImageList, subItem.iImage, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, ILD_MASK | (subItem.state & LVIS_OVERLAYMASK));
					}
					// TODO: Don't use hard-coded margins
					subItemRect.left = iconBoundingRect.right;
				}

				// draw the text
				if(layoutRTL) {
					// TODO: Don't use hard-coded margins
					subItemRect.OffsetRect(1, 0);
				}
				// TODO: Don't use hard-coded margins
				subItemRect.InflateRect(-2, 0);
				switch(column.fmt & LVCFMT_JUSTIFYMASK) {
					case LVCFMT_LEFT:
						// TODO: Don't use hard-coded margins
						subItemRect.left += 4;
						break;
					case LVCFMT_RIGHT:
						// TODO: Don't use hard-coded margins
						subItemRect.right -= 4;
						break;
				}
				if(themedListItems) {
					CT2W converter(subItem.pszText);
					LPWSTR pLabelText = converter;
					themingEngine.DrawThemeText(memoryDC, LVP_LISTITEM, themeState, pLabelText, lstrlenW(pLabelText), textDrawStyle, 0, &subItemRect);
				} else {
					memoryDC.DrawText(subItem.pszText, lstrlen(subItem.pszText), &subItemRect, textDrawStyle);
				}
				if(!itemIsSelected) {
					COLORREF bkColor = memoryDC.GetBkColor();
					for(int y = subItemRect.top; y <= subItemRect.bottom; ++y) {
						for(int x = subItemRect.left; x <= subItemRect.right; ++x) {
							if(memoryDC.GetPixel(x, y) != bkColor) {
								maskMemoryDC.SetPixelV(x, y, 0x00000000);
							}
						}
					}
				}
			}
			if(hMod) {
				FreeLibrary(hMod);
			}
		}
	}

	if(!themedListItems) {
		// draw the focus rectangle
		if(hasFocus && itemIsFocused) {
			if((SendMessage(WM_QUERYUISTATE, 0, 0) & UISF_HIDEFOCUS) == 0) {
				memoryDC.DrawFocusRect(&focusRect);
				maskMemoryDC.FrameRect(&focusRect, static_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
			}
		}
	}

	if(doAlphaChannelProcessing) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						if(themedListItems) {
							if(itemIsSelected) {
								if((pPixel->rgbReserved != 0x00) || selectionBoundingRect.PtInRect(pt2)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else if(labelBoundingRect.PtInRect(pt2)) {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						} else {
							// items are not themed
							if(itemIsSelected) {
								if((pPixel->rgbReserved == 0x00) || selectionBoundingRect.PtInRect(pt2)) {
									if(subItemImages) {
										BOOL changeAlpha = TRUE;
										for(int i = 0; i < columnCount - 1; ++i) {
											if(PtInRect(&pSubItemImageRectangles[i], pt2)) {
												changeAlpha = !pSubItemImageHasAlpha[i];
												break;
											}
										}
										if(changeAlpha) {
											pPixel->rgbReserved = 0xFF;
										}
									} else {
										pPixel->rgbReserved = 0xFF;
									}
								}
							} else {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(themedListItems) {
							if(itemIsSelected) {
								if((pPixel->rgbReserved != 0x00) || selectionBoundingRect.PtInRect(pt)) {
									pPixel->rgbReserved = 0xFF;
								}
							} else if(labelBoundingRect.PtInRect(pt)) {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						} else {
							// items are not themed
							if(itemIsSelected) {
								if((pPixel->rgbReserved == 0x00) || selectionBoundingRect.PtInRect(pt)) {
									if(subItemImages) {
										BOOL changeAlpha = TRUE;
										for(int i = 0; i < columnCount - 1; ++i) {
											if(PtInRect(&pSubItemImageRectangles[i], pt)) {
												changeAlpha = !pSubItemImageHasAlpha[i];
												break;
											}
										}
										if(changeAlpha) {
											pPixel->rgbReserved = 0xFF;
										}
									} else {
										pPixel->rgbReserved = 0xFF;
									}
								}
							} else {
								if(pPixel->rgbReserved == 0x00) {
									pPixel->rgbReserved = 0xFF;
								}
							}
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
	if(hPreviousFont) {
		memoryDC.SelectFont(hPreviousFont);
	}

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	VariantClear(&vTileViewColumns);
	ReleaseDC(hCompatibleDC);
	if(pSubItemImageRectangles) {
		delete[] pSubItemImageRectangles;
	}
	if(pSubItemImageHasAlpha) {
		delete[] pSubItemImageHasAlpha;
	}

	return hDragImageList;
}

HIMAGELIST ExplorerListView::CreateLegacyHeaderDragImage(int columnIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	if(pBoundingRectangle) {
		containedSysHeader32.SendMessage(HDM_GETITEMRECT, columnIndex, reinterpret_cast<LPARAM>(pBoundingRectangle));
	}
	if(pUpperLeftPoint) {
		if(pBoundingRectangle) {
			pUpperLeftPoint->x = pBoundingRectangle->left;
			pUpperLeftPoint->y = pBoundingRectangle->top;
		} else {
			RECT rc = {0};
			if(containedSysHeader32.SendMessage(HDM_GETITEMRECT, columnIndex, reinterpret_cast<LPARAM>(&rc))) {
				pUpperLeftPoint->x = rc.left;
				pUpperLeftPoint->y = rc.top;
			}
		}
	}

	return reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_CREATEDRAGIMAGE, columnIndex, 0));
}

BOOL ExplorerListView::CreateLegacyOLEDragImage(IListViewItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pItems);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	OLE_HANDLE h = NULL;
	OLE_XPOS_PIXELS xUpperLeft = 0;
	OLE_YPOS_PIXELS yUpperLeft = 0;
	pItems->CreateDragImage(&xUpperLeft, &yUpperLeft, &h);
	if(h) {
		HIMAGELIST hImageList = static_cast<HIMAGELIST>(LongToHandle(h));

		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD flags = 0;
			pImgLst->GetItemFlags(0, &flags);
			if(flags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				pDragImage->crColorKey = 0xFFFFFFFF;

				// transfer raw data
				CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(NULL);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(NULL, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(WTL::CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = xUpperLeft + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - xUpperLeft;
			}
			pDragImage->ptOffset.y = mousePosition.y - yUpperLeft;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

BOOL ExplorerListView::CreateLegacyOLEHeaderDragImage(IListViewColumn* pColumn, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pColumn);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	OLE_HANDLE h = NULL;
	OLE_XPOS_PIXELS xUpperLeft = 0;
	OLE_YPOS_PIXELS yUpperLeft = 0;
	pColumn->CreateDragImage(&xUpperLeft, &yUpperLeft, &h);
	if(h) {
		HIMAGELIST hImageList = static_cast<HIMAGELIST>(LongToHandle(h));

		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD flags = 0;
			pImgLst->GetItemFlags(0, &flags);
			if(flags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				pDragImage->crColorKey = 0xFFFFFFFF;

				// transfer raw data
				CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(NULL);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(NULL, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(WTL::CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			containedSysHeader32.ScreenToClient(&mousePosition);
			if(containedSysHeader32.GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = xUpperLeft + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - xUpperLeft;
			}
			pDragImage->ptOffset.y = mousePosition.y - yUpperLeft;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

BOOL ExplorerListView::CreateVistaOLEDragImage(IListViewItemContainer* pItems, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pItems);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	CTheme themingEngine;
	themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
	if(themingEngine.IsThemeNull()) {
		// FIXME: What should we do here?
		ATLASSERT(FALSE && "Current theme does not define the \"DragDrop\" class.");
	} else {
		// retrieve the drag image's size
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();

		themingEngine.GetThemePartSize(memoryDC, DD_IMAGEBG, 1, NULL, TS_TRUE, &pDragImage->sizeDragImage);
		MARGINS margins = {0};
		themingEngine.GetThemeMargins(memoryDC, DD_IMAGEBG, 1, TMT_CONTENTMARGINS, NULL, &margins);
		pDragImage->sizeDragImage.cx -= margins.cxLeftWidth + margins.cxRightWidth;
		pDragImage->sizeDragImage.cy -= margins.cyTopHeight + margins.cyBottomHeight;
	}

	ATLASSERT(pDragImage->sizeDragImage.cx > 0);
	ATLASSERT(pDragImage->sizeDragImage.cy > 0);

	// create target bitmap
	BITMAPINFO bitmapInfo = {0};
	bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
	bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
	bitmapInfo.bmiHeader.biPlanes = 1;
	bitmapInfo.bmiHeader.biBitCount = 32;
	bitmapInfo.bmiHeader.biCompression = BI_RGB;
	LPRGBQUAD pDragImageBits = NULL;
	pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);

	HIMAGELIST hSourceImageList = (cachedSettings.hHighResImageList ? cachedSettings.hHighResImageList : (cachedSettings.hExtraLargeImageList ? cachedSettings.hExtraLargeImageList : (cachedSettings.hLargeImageList ? cachedSettings.hLargeImageList : cachedSettings.hSmallImageList)));
	if(!hSourceImageList) {
		// report success, although we've an empty drag image
		return TRUE;
	}

	IImageList2* pImgLst = NULL;
	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
		HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
		if(pfnHIMAGELIST_QueryInterface) {
			pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
		}
		FreeLibrary(hMod);
	}
	if(!pImgLst) {
		IImageList* p = reinterpret_cast<IImageList*>(hSourceImageList);
		p->QueryInterface(IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
	}
	ATLASSUME(pImgLst);

	if(pImgLst) {
		LONG numberOfItems = 0;
		pItems->Count(&numberOfItems);
		ATLASSERT(numberOfItems > 0);
		// don't display more than 5 (10) thumbnails
		numberOfItems = min(numberOfItems, (hSourceImageList == cachedSettings.hHighResImageList ? 5 : 10));

		CComPtr<IUnknown> pUnknownEnum = NULL;
		pItems->get__NewEnum(&pUnknownEnum);
		CComQIPtr<IEnumVARIANT> pEnum = pUnknownEnum;
		ATLASSUME(pEnum);
		if(pEnum) {
			int cx = 0;
			int cy = 0;
			pImgLst->GetIconSize(&cx, &cy);
			SIZE thumbnailSize;
			thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
			if(thumbnailSize.cy < 8) {
				// don't get smaller than 8x8 thumbnails
				numberOfItems = (pDragImage->sizeDragImage.cy - 8) / 3 + 1;
				thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
			}
			thumbnailSize.cx = thumbnailSize.cy;
			int thumbnailBufferSize = thumbnailSize.cx * thumbnailSize.cy * sizeof(RGBQUAD);
			LPRGBQUAD pThumbnailBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, thumbnailBufferSize));
			ATLASSERT(pThumbnailBits);
			if(pThumbnailBits) {
				// iterate over the dragged items
				VARIANT v;
				int i = 0;
				CComPtr<IListViewItem> pItem = NULL;
				while(pEnum->Next(1, &v, NULL) == S_OK) {
					if(v.vt == VT_DISPATCH) {
						v.pdispVal->QueryInterface(IID_IListViewItem, reinterpret_cast<LPVOID*>(&pItem));
						ATLASSUME(pItem);
						if(pItem) {
							// get the item's icon
							LONG icon = 0;
							pItem->get_IconIndex(&icon);

							pImgLst->ForceImagePresent(icon, ILFIP_ALWAYS);
							HICON hIcon = NULL;
							pImgLst->GetIcon(icon, ILD_TRANSPARENT, &hIcon);
							ATLASSERT(hIcon);
							if(hIcon) {
								// finally create the thumbnail
								ZeroMemory(pThumbnailBits, thumbnailBufferSize);
								HRESULT hr = CreateThumbnail(hIcon, thumbnailSize, pThumbnailBits, TRUE);
								DestroyIcon(hIcon);
								if(FAILED(hr)) {
									pItem = NULL;
									VariantClear(&v);
									break;
								}

								// add the thumbail to the drag image keeping the alpha channel intact
								if(i == 0) {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx, pThumbnailPixel += thumbnailSize.cx) {
										CopyMemory(pDragImagePixel, pThumbnailPixel, thumbnailSize.cx * sizeof(RGBQUAD));
									}
								} else {
									LPRGBQUAD pDragImagePixel = pDragImageBits;
									LPRGBQUAD pThumbnailPixel = pThumbnailBits;
									pDragImagePixel += 3 * i * pDragImage->sizeDragImage.cx;
									for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx) {
										LPRGBQUAD p = pDragImagePixel + 2 * i;
										for(int x = 0; x < thumbnailSize.cx; ++x, ++p, ++pThumbnailPixel) {
											// merge the pixels
											p->rgbRed = pThumbnailPixel->rgbRed * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbRed / 0xFF;
											p->rgbGreen = pThumbnailPixel->rgbGreen * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbGreen / 0xFF;
											p->rgbBlue = pThumbnailPixel->rgbBlue * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbBlue / 0xFF;
											p->rgbReserved = pThumbnailPixel->rgbReserved + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbReserved / 0xFF;
										}
									}
								}
							}

							++i;
							pItem = NULL;
							if(i == numberOfItems) {
								VariantClear(&v);
								break;
							}
						}
					}
					VariantClear(&v);
				}
				HeapFree(GetProcessHeap(), 0, pThumbnailBits);
				succeeded = TRUE;
			}
		}

		pImgLst->Release();
	}

	return succeeded;
}

BOOL ExplorerListView::CreateVistaOLEHeaderDragImage(IListViewColumn* pColumn, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pColumn);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;
	HIMAGELIST hSourceImageList = reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_NORMAL, 0));

	CTheme themingEngine;
	themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
	if(themingEngine.IsThemeNull()) {
		// FIXME: What should we do here?
		ATLASSERT(FALSE && "Current theme does not define the \"DragDrop\" class.");
	} else {
		// retrieve the drag image's size
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();

		themingEngine.GetThemePartSize(memoryDC, DD_IMAGEBG, 1, NULL, TS_TRUE, &pDragImage->sizeDragImage);
		MARGINS margins = {0};
		themingEngine.GetThemeMargins(memoryDC, DD_IMAGEBG, 1, TMT_CONTENTMARGINS, NULL, &margins);
		pDragImage->sizeDragImage.cx -= margins.cxLeftWidth + margins.cxRightWidth;
		pDragImage->sizeDragImage.cy -= margins.cyTopHeight + margins.cyBottomHeight;
	}

	ATLASSERT(pDragImage->sizeDragImage.cx > 0);
	ATLASSERT(pDragImage->sizeDragImage.cy > 0);

	// create target bitmap
	BITMAPINFO bitmapInfo = {0};
	bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
	bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
	bitmapInfo.bmiHeader.biPlanes = 1;
	bitmapInfo.bmiHeader.biBitCount = 32;
	bitmapInfo.bmiHeader.biCompression = BI_RGB;
	LPRGBQUAD pDragImageBits = NULL;
	pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);

	hSourceImageList = (cachedSettings.hHighResHeaderImageList ? cachedSettings.hHighResHeaderImageList : hSourceImageList);
	if(!hSourceImageList) {
		// report success, although we've an empty drag image
		return TRUE;
	}

	IImageList2* pImgLst = NULL;
	HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
	if(hMod) {
		typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
		HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
		if(pfnHIMAGELIST_QueryInterface) {
			pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
		}
		FreeLibrary(hMod);
	}
	if(!pImgLst) {
		IImageList* p = reinterpret_cast<IImageList*>(hSourceImageList);
		p->QueryInterface(IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
	}
	ATLASSUME(pImgLst);

	if(pImgLst) {
		LONG numberOfItems = 1;

		int cx = 0;
		int cy = 0;
		pImgLst->GetIconSize(&cx, &cy);
		SIZE thumbnailSize;
		thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfItems - 1);
		thumbnailSize.cx = thumbnailSize.cy;
		int thumbnailBufferSize = thumbnailSize.cx * thumbnailSize.cy * sizeof(RGBQUAD);
		LPRGBQUAD pThumbnailBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, thumbnailBufferSize));
		ATLASSERT(pThumbnailBits);
		if(pThumbnailBits) {
			if(pColumn) {
				// don't get smaller than 8x8 thumbnails
				LONG icon = 0;
				pColumn->get_IconIndex(&icon);

				pImgLst->ForceImagePresent(icon, ILFIP_ALWAYS);
				HICON hIcon = NULL;
				pImgLst->GetIcon(icon, ILD_TRANSPARENT, &hIcon);
				ATLASSERT(hIcon);
				if(hIcon) {
					// finally create the thumbnail
					ZeroMemory(pThumbnailBits, thumbnailBufferSize);
					ATLVERIFY(SUCCEEDED(CreateThumbnail(hIcon, thumbnailSize, pThumbnailBits, TRUE)));
					DestroyIcon(hIcon);

					// add the thumbail to the drag image keeping the alpha channel intact
					LPRGBQUAD pDragImagePixel = pDragImageBits;
					LPRGBQUAD pThumbnailPixel = pThumbnailBits;
					for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx, pThumbnailPixel += thumbnailSize.cx) {
						CopyMemory(pDragImagePixel, pThumbnailPixel, thumbnailSize.cx * sizeof(RGBQUAD));
					}
				}
			}
			HeapFree(GetProcessHeap(), 0, pThumbnailBits);
			succeeded = TRUE;
		}

		pImgLst->Release();
	}

	return succeeded;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ExplorerListView::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	HWND hWndToUse = NULL;
	if(WindowFromPoint(buffer) == containedSysHeader32) {
		hWndToUse = containedSysHeader32;
		Raise_HeaderOLEDragEnter(FALSE, pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	} else {
		hWndToUse = *this;
		Raise_OLEDragEnter(FALSE, pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	}

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(hWndToUse, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(hWndToUse, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::DragLeave(void)
{
	KillTimer(timers.ID_DRAGSCROLL);
	BOOL callDropTargetHelper = TRUE;
	if(dragDropStatus.isOverHeader) {
		Raise_HeaderOLEDragLeave(FALSE, &callDropTargetHelper);
	} else {
		Raise_OLEDragLeave(FALSE, &callDropTargetHelper);
	}

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragLeave();
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;

	if(WindowFromPoint(buffer) == containedSysHeader32) {
		if(dragDropStatus.isOverHeader) {
			Raise_HeaderOLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);
		} else {
			// we've left the listview and entered the header control
			Raise_OLEDragLeave(TRUE, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_HeaderOLEDragEnter(TRUE, NULL, pEffect, keyState, mousePosition, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(containedSysHeader32, pDataObject, &buffer, *pEffect);
			}
		}
	} else {
		if(dragDropStatus.isOverHeader) {
			// we've left the header control and entered the listview
			Raise_HeaderOLEDragLeave(TRUE, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragLeave();
			}

			DROPDESCRIPTION newDropDescription;
			ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
			if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
				InvalidateDragWindow(pDataObject);
				oldDropDescription = newDropDescription;
			}

			Raise_OLEDragEnter(TRUE, NULL, pEffect, keyState, mousePosition, &callDropTargetHelper);
			if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
				dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			}
		} else {
			Raise_OLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);
		}
	}

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	BOOL callDropTargetHelper = TRUE;
	if(WindowFromPoint(buffer) == containedSysHeader32) {
		Raise_HeaderOLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	} else {
		Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);
	}

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP ExplorerListView::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL useDropDescriptionHack = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				useDropDescriptionHack = *reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(useDropDescriptionHack) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *reinterpret_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	if(dragDropStatus.headerIsSource) {
		Raise_HeaderOLEGiveFeedback(effect, &useDefaultCursors);
	} else {
		Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	}
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP ExplorerListView::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	if(dragDropStatus.headerIsSource) {
		Raise_HeaderOLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	} else {
		Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	}
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP ExplorerListView::DragEnterTarget(HWND hWndTarget)
{
	if(dragDropStatus.headerIsSource) {
		Raise_HeaderOLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	} else {
		Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::DragLeaveTarget(void)
{
	if(dragDropStatus.headerIsSource) {
		Raise_HeaderOLEDragLeavePotentialTarget();
	} else {
		Raise_OLEDragLeavePotentialTarget();
	}
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IListViewFooterCallback
STDMETHODIMP ExplorerListView::OnButtonClicked(int itemIndex, LPARAM /*lParam*/, PINT pRemoveFooter)
{
	CComPtr<IListViewFooterItem> pLvwFooterItem = ClassFactory::InitFooterItem(itemIndex, this, FALSE);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	//button |= 1/*MouseButtonConstants.vbLeftButton*/;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, NULL, NULL, TRUE, TRUE, FALSE);

	VARIANT_BOOL removeFooterArea = BOOL2VARIANTBOOL(*pRemoveFooter);
	Raise_FooterItemClick(pLvwFooterItem, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &removeFooterArea);
	*pRemoveFooter = VARIANTBOOL2BOOL(removeFooterArea);
	return S_OK;
}

STDMETHODIMP ExplorerListView::OnDestroyButton(int itemIndex, LPARAM lParam)
{
	if(!(properties.disabledEvents & deFreeFooterItemData)) {
		CComPtr<IListViewFooterItem> pLvwFooterItem = ClassFactory::InitFooterItem(itemIndex, this, FALSE);
		Raise_FreeFooterItemData(pLvwFooterItem, static_cast<LONG>(lParam));
	}
	return S_OK;
}
// implementation of IListViewFooterCallback
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOwnerDataCallback
STDMETHODIMP ExplorerListView::GetItemPosition(int /*itemIndex*/, LPPOINT /*pPosition*/)
{
	ATLASSERT(FALSE && "GetItemPosition");
	return E_NOTIMPL;
}

STDMETHODIMP ExplorerListView::SetItemPosition(int /*itemIndex*/, POINT /*position*/)
{
	ATLASSERT(FALSE && "SetItemPosition");
	return E_NOTIMPL;
}

STDMETHODIMP ExplorerListView::GetItemInGroup(int groupIndex, int groupWideItemIndex, PINT pTotalItemIndex)
{
	Raise_MapGroupWideToTotalItemIndex(groupIndex, groupWideItemIndex, reinterpret_cast<LONG*>(pTotalItemIndex));
	return S_OK;
}

STDMETHODIMP ExplorerListView::GetItemGroup(int itemIndex, int occurenceIndex, PINT pGroupIndex)
{
	Raise_ItemGetGroup(itemIndex, occurenceIndex, reinterpret_cast<LONG*>(pGroupIndex));
	return S_OK;
}

STDMETHODIMP ExplorerListView::GetItemGroupCount(int itemIndex, PINT pOccurenceCount)
{
	*pOccurenceCount = 1;
	Raise_ItemGetOccurrencesCount(itemIndex, reinterpret_cast<LONG*>(pOccurenceCount));
	return S_OK;
}

STDMETHODIMP ExplorerListView::OnCacheHint(LVITEMINDEX firstItem, LVITEMINDEX lastItem)
{
	CComPtr<IListViewItem> pFirstItem = ClassFactory::InitListItem(firstItem, this, FALSE);
	CComPtr<IListViewItem> pLastItem = ClassFactory::InitListItem(lastItem, this, FALSE);
	Raise_CacheItemsHint(pFirstItem, pLastItem);
	return S_OK;
}
// implementation of IOwnerDataCallback
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ISubItemCallback
STDMETHODIMP ExplorerListView::GetSubItemTitle(int subItemIndex, LPWSTR pBuffer, int bufferSize)
{
	if(!pBuffer) {
		return E_POINTER;
	}

	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		HDITEMW column = {0};
		column.cchTextMax = bufferSize;
		column.pszText = pBuffer;
		column.mask = HDI_TEXT;
		if(containedSysHeader32.SendMessage(HDM_GETITEM, subItemIndex, reinterpret_cast<LPARAM>(&column))) {
			return S_OK;
		}
		return E_INVALIDARG;
	#else
		return E_NOTIMPL;
	#endif
}

STDMETHODIMP ExplorerListView::GetSubItemControl(int itemIndex, int subItemIndex, REFIID requiredInterface, LPVOID* ppObject)
{
	ATLASSUME(ppObject);
	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		/* NOTE: We pass -2 here, because it wouldn't be suprising if the native list view control would set
		 *       mode to VARIANT_TRUE in certain situations. */
		return BeginSubItemEdit(itemIndex, subItemIndex, static_cast<int>(siemNotSpecified), requiredInterface, ppObject);
	#else
		return E_NOTIMPL;
	#endif
}

STDMETHODIMP ExplorerListView::BeginSubItemEdit(int itemIndex, int subItemIndex, int mode, REFIID requiredInterface, LPVOID* ppObject)
{
	if(!ppObject) {
		return E_POINTER;
	}
	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(!shellBrowserInterface.pInternalMessageListener) {
				if(properties.disabledEvents & deGetSubItemControl) {
					return E_NOINTERFACE;
				}
			}
		#else
			if(properties.disabledEvents & deGetSubItemControl) {
				return E_NOINTERFACE;
			}
		#endif

		SubItemControlKindConstants controlKind;
		if(requiredInterface == IID_IDrawPropertyControl || requiredInterface == IID_IWin10Unknown) {
			controlKind = sickVisualRepresentation;
		} else if(requiredInterface == IID_IPropertyControl) {
			controlKind = sickInPlaceEditing;
		} else {
			ATLASSERT(FALSE && "BeginSubItemEdit: Unknown IID");
			return E_NOINTERFACE;
		}
		SubItemControlConstants subItemControl = sicNoSubItemControl;
		HRESULT hr = E_NOINTERFACE;
		IPropertyDescription* pPropertyDescription = NULL;
		PROPVARIANT propertyValue = {0};
		PropVariantInit(&propertyValue);
		PROPVARIANT* pPropertyValue = &propertyValue;
		#ifdef INCLUDESHELLBROWSERINTERFACE
			SHLVMGETSUBITEMCONTROLDATA subItemControlData = {0};
			if(shellBrowserInterface.pInternalMessageListener) {
				subItemControlData.structSize = sizeof(subItemControlData);
				subItemControlData.itemID = ItemIndexToID(itemIndex);
				subItemControlData.columnID = ColumnIndexToID(subItemIndex);
				subItemControlData.controlKind = controlKind;
				subItemControlData.subItemEditMode = mode;
				subItemControlData.subItemControl = static_cast<int>(subItemControl);

				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_GETSUBITEMCONTROL, 0, reinterpret_cast<LPARAM>(&subItemControlData));

				if(subItemControlData.useSubItemControl) {
					subItemControl = static_cast<SubItemControlConstants>(subItemControlData.subItemControl);
					pPropertyDescription = subItemControlData.pPropertyDescription;
					pPropertyValue = &subItemControlData.propertyValue;
				} else if(subItemControlData.pPropertyDescription) {
					subItemControlData.pPropertyDescription->Release();
					subItemControlData.pPropertyDescription = NULL;
					if(properties.disabledEvents & deGetSubItemControl) {
						return E_NOINTERFACE;
					}
				}
			}
		#endif
		LVITEMINDEX parentItemIndex = {itemIndex, 0};
		CComPtr<IListViewSubItem> pSubItem = ClassFactory::InitListSubItem(parentItemIndex, subItemIndex, this);

		if(!(properties.disabledEvents & deGetSubItemControl)) {
			Raise_GetSubItemControl(pSubItem, controlKind, static_cast<SubItemEditModeConstants>(mode), &subItemControl);
		}

		switch(subItemControl) {
			/*case sicUseEvents:
				// TODO: create instance of custom implementation
				break;*/
			case sicMultiLineText:
				hr = CoCreateInstance(CLSID_CInPlaceMLEditBoxControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				if(!pPropertyDescription) {
					PSGetPropertyDescriptionByName(L"System.Generic.String", IID_PPV_ARGS(&pPropertyDescription));
				}
				break;
			// TODO: ShellListView uses this value!
			/*case sicMultiValueText:
				if(requiredInterface == IID_IDrawPropertyControl || requiredInterface == IID_IWin10Unknown) {
					hr = CoCreateInstance(CLSID_CCustomDrawMultiValuePropertyControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				} else {
					hr = CoCreateInstance(CLSID_CInPlaceMultiValuePropertyControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				}
				if(!pPropertyDescription) {
					PSGetPropertyDescriptionByName(L"System.Keywords", IID_PPV_ARGS(&pPropertyDescription));
				}
				break;*/
			case sicPercentBar:
				hr = CoCreateInstance(CLSID_CCustomDrawPercentFullControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				break;
			/*case sicProgressBar:
				hr = CoCreateInstance(CLSID_CCustomDrawProgressControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				break;*/
			case sicRating:
				hr = CoCreateInstance(CLSID_CRatingControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				break;
			case sicText:
				if(requiredInterface == IID_IDrawPropertyControl || requiredInterface == IID_IWin10Unknown) {
					hr = CoCreateInstance(CLSID_CStaticPropertyControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				} else {
					hr = CoCreateInstance(CLSID_CInPlaceEditBoxControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				}
				if(!pPropertyDescription) {
					PSGetPropertyDescriptionByName(L"System.Generic.String", IID_PPV_ARGS(&pPropertyDescription));
				}
				break;
			/*case sicIconList:
				hr = CoCreateInstance(CLSID_CIconListControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				break;*/
			case sicBooleanCheckMark:
			case sicCheckboxDropList:
				hr = CoCreateInstance(CLSID_CBooleanControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				if(!pPropertyDescription) {
					PSGetPropertyDescriptionByName(L"System.Generic.Boolean", IID_PPV_ARGS(&pPropertyDescription));
				}
				break;
			case sicCalendar:
				hr = CoCreateInstance(CLSID_CInPlaceCalendarControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				if(!pPropertyDescription) {
					PSGetPropertyDescriptionByName(L"System.Generic.DateTime", IID_PPV_ARGS(&pPropertyDescription));
				}
				break;
			case sicDropList:
				hr = CoCreateInstance(CLSID_CInPlaceDropListComboControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				break;
			case sicHyperlink:
				hr = CoCreateInstance(CLSID_CHyperlinkControl, NULL, CLSCTX_INPROC_SERVER, requiredInterface, ppObject);
				break;
		}
		if(SUCCEEDED(hr)) {
			IPropertyControlBase* pControl = *reinterpret_cast<IPropertyControlBase**>(ppObject);
			CComBSTR themeAppName = L"";
			CComBSTR themeIDList = L"";
			if(CTheme::IsThemingSupported()) {
				WCHAR pSubAppNameBuffer[300] = {0};
				LPWSTR pSubAppName = pSubAppNameBuffer;
				ATOM valueSubAppName = reinterpret_cast<ATOM>(GetPropW(*this, L"#43281"));
				if(valueSubAppName) {
					GetAtomNameW(valueSubAppName, pSubAppNameBuffer, 300);
					if(lstrlenW(pSubAppNameBuffer) == 1 && pSubAppNameBuffer[0] == L'$') {
						pSubAppNameBuffer[0] = L'\0';
					}
				} else {
					pSubAppName = NULL;
				}
				themeAppName = pSubAppName;
				WCHAR pSubIDListBuffer[300] = {0};
				LPWSTR pSubIDList = pSubIDListBuffer;
				ATOM valueSubIDList = reinterpret_cast<ATOM>(GetPropW(*this, L"#43280"));
				if(valueSubIDList) {
					GetAtomNameW(valueSubIDList, pSubIDListBuffer, 300);
					if(lstrlenW(pSubIDListBuffer) == 1 && pSubIDListBuffer[0] == L'$') {
						pSubIDListBuffer[0] = L'\0';
					}
				} else {
					pSubIDList = NULL;
				}
				themeIDList = pSubIDList;
			}
			HFONT hFont = GetFont();
			COLORREF textColor = static_cast<COLORREF>(SendMessage(LVM_GETTEXTCOLOR, 0, 0));
			if(textColor == CLR_NONE) {
				textColor = GetSysColor(COLOR_WINDOWTEXT);
			}

			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(!subItemControlData.useSubItemControl || !subItemControlData.hasSetValue) {
			#endif
			LPWSTR pBuffer = reinterpret_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (MAX_ITEMTEXTLENGTH + 1) * sizeof(WCHAR)));
			if(pBuffer) {
				LVITEMW item = {0};
				item.iSubItem = subItemIndex;
				item.cchTextMax = MAX_ITEMTEXTLENGTH;
				item.pszText = pBuffer;
				SendMessage(LVM_GETITEMTEXTW, parentItemIndex.iItem, reinterpret_cast<LPARAM>(&item));
				if(subItemControl == sicPercentBar || subItemControl == sicRating) {
					PROPVARIANT tmp;
					PropVariantInit(&tmp);
					InitPropVariantFromString(item.pszText, &tmp);
					// NOTE: Converting to VT_INT instead of VT_UI4 will work on Windows 7, but not on Vista.
					PropVariantChangeType(pPropertyValue, tmp, 0, VT_UI4);
					PropVariantClear(&tmp);
				} else if(subItemControl == sicBooleanCheckMark || subItemControl == sicCheckboxDropList) {
					PROPVARIANT tmp;
					PropVariantInit(&tmp);
					InitPropVariantFromString(item.pszText, &tmp);
					PropVariantChangeType(pPropertyValue, tmp, 0, VT_BOOL);
					PropVariantClear(&tmp);
				} else if(subItemControl == sicCalendar) {
					// problem 1: VT_FILETIME is the prefered format, but VT_FILETIME isn't very VB6-friendly, so convert using VT_DATE
					// problem 2: VT_FILETIME is a UTC-based time, but VT_DATE is treated as a local time
					VARIANT v;
					VariantInit(&v);
					v.bstrVal = bstr_t(item.pszText).Detach();
					v.vt = VT_BSTR;
					VariantChangeType(&v, &v, 0, VT_DATE);
					PROPVARIANT tmp;
					PropVariantInit(&tmp);
					VariantToPropVariant(&v, &tmp);
					VariantClear(&v);
					PropVariantChangeType(pPropertyValue, tmp, 0, VT_FILETIME);
					PropVariantClear(&tmp);
					FILETIME ft = pPropertyValue->filetime;
					LocalFileTimeToFileTime(&ft, &pPropertyValue->filetime);
				} else {
					InitPropVariantFromString(item.pszText, pPropertyValue);
				}
				HeapFree(GetProcessHeap(), 0, pBuffer);
			}
			#ifdef INCLUDESHELLBROWSERINTERFACE
				}
			#endif

			LPUNKNOWN pPropertyDescriptionUnknown = NULL;
			if(pPropertyDescription) {
				pPropertyDescription->QueryInterface(IID_PPV_ARGS(&pPropertyDescriptionUnknown));
			}
			if(!(properties.disabledEvents & deGetSubItemControl)) {
				Raise_ConfigureSubItemControl(pSubItem, controlKind, static_cast<SubItemEditModeConstants>(mode), subItemControl, &themeAppName, &themeIDList, &hFont, &textColor, &pPropertyDescriptionUnknown, pPropertyValue);
			}
			if(pPropertyDescriptionUnknown) {
				if(pPropertyDescription) {
					pPropertyDescription->Release();
					pPropertyDescription = NULL;
				}
				pPropertyDescriptionUnknown->QueryInterface(IID_PPV_ARGS(&pPropertyDescription));
				pPropertyDescriptionUnknown->Release();
				pPropertyDescriptionUnknown = NULL;
			} else if(pPropertyDescription) {
				pPropertyDescription->Release();
				pPropertyDescription = NULL;
			}

			CComPtr<IPropertyValue> pPropertyValueObj = NULL;
			if(SUCCEEDED(IPropertyValueImpl::CreateInstance(NULL, IID_IPropertyValue, reinterpret_cast<LPVOID*>(&pPropertyValueObj))) && pPropertyValueObj && SUCCEEDED(pPropertyValueObj->InitValue(*pPropertyValue))) {
				if(pPropertyDescription) {
					ATLVERIFY(SUCCEEDED(pControl->Initialize(pPropertyDescription, static_cast<IPropertyControlBase::PROPDESC_CONTROL_TYPE>(0))));
					ATLVERIFY(SUCCEEDED(pControl->SetValue(pPropertyValueObj)));
				} else {
					pControl->SetValue(pPropertyValueObj);
				}
				ATLVERIFY(SUCCEEDED(pControl->SetTextColor(textColor)));
				if(hFont) {
					ATLVERIFY(SUCCEEDED(pControl->SetFont(hFont)));
				}
				ATLVERIFY(SUCCEEDED(pControl->SetWindowTheme(themeAppName, themeIDList)));
			} else {
				pControl->Destroy();
				hr = E_NOINTERFACE;
			}

			if(pPropertyDescription) {
				pPropertyDescription->Release();
				pPropertyDescription = NULL;
			}
		}

		return hr;
	#else
		return E_NOTIMPL;
	#endif
}

STDMETHODIMP ExplorerListView::EndSubItemEdit(int itemIndex, int subItemIndex, int mode, IPropertyControl* pPropertyControl)
{
	ATLASSUME(pPropertyControl);
	if(!pPropertyControl) {
		return E_POINTER;
	}

	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		BOOL modified = FALSE;
		pPropertyControl->IsModified(&modified);
		if(modified) {
			CComPtr<IPropertyValue> pPropertyValue = NULL;
			if(SUCCEEDED(pPropertyControl->GetValue(IID_IPropertyValue, reinterpret_cast<LPVOID*>(&pPropertyValue))) && pPropertyValue) {
				#ifdef INCLUDESHELLBROWSERINTERFACE
					SHLVMGETSUBITEMCONTROLDATA subItemControlData = {0};
					LPPROPVARIANT pPropValue = &subItemControlData.propertyValue;
				#else
					PROPVARIANT propertyValue;
					LPPROPVARIANT pPropValue = &propertyValue;
				#endif
				PropVariantInit(pPropValue);
				if(SUCCEEDED(pPropertyValue->GetValue(pPropValue))) {
					PROPERTYKEY propertyKey = {0};
					pPropertyValue->GetPropertyKey(&propertyKey);

					DATE originalDateValue = 0;
					FILETIME originalFileTimeValue = {0};
					VARTYPE originalVarType = pPropValue->vt;
					if(pPropValue->vt == VT_FILETIME) {
						// problem 1: VT_FILETIME isn't very VB6-friendly, so convert it to VT_DATE
						// problem 2: VT_FILETIME is a UTC-based time, but VT_DATE is treated as a local time
						originalFileTimeValue = pPropValue->filetime;
						FILETIME ft = pPropValue->filetime;
						FileTimeToLocalFileTime(&ft, &pPropValue->filetime);
						PROPVARIANT tmp;
						PropVariantInit(&tmp);
						PropVariantCopy(&tmp, pPropValue);
						PropVariantClear(pPropValue);
						PropVariantChangeType(pPropValue, tmp, 0, VT_DATE);
						originalDateValue = pPropValue->date;
					}

					LVITEMINDEX parentItemIndex = {itemIndex, 0};
					CComPtr<IListViewSubItem> pSubItem = ClassFactory::InitListSubItem(parentItemIndex, subItemIndex, this);

					VARIANT_BOOL cancel = VARIANT_FALSE;
					Raise_EndSubItemEdit(pSubItem, static_cast<SubItemEditModeConstants>(mode), &propertyKey, pPropValue, &cancel);

					#ifdef INCLUDESHELLBROWSERINTERFACE
						if(shellBrowserInterface.pInternalMessageListener && cancel == VARIANT_FALSE) {
							if(originalVarType == VT_FILETIME && pPropValue->vt == VT_DATE) {
								if(originalDateValue == pPropValue->date) {
									// restore original VT_FILETIME value
									PropVariantClear(pPropValue);
									pPropValue->vt = VT_FILETIME;
									pPropValue->filetime = originalFileTimeValue;
								} else {
									// convert back to VT_FILETIME
									PROPVARIANT tmp;
									PropVariantInit(&tmp);
									PropVariantCopy(&tmp, pPropValue);
									PropVariantClear(pPropValue);
									PropVariantChangeType(pPropValue, tmp, 0, VT_FILETIME);
									FILETIME ft = pPropValue->filetime;
									LocalFileTimeToFileTime(&ft, &pPropValue->filetime);
								}
							}

							subItemControlData.structSize = sizeof(subItemControlData);
							subItemControlData.itemID = ItemIndexToID(itemIndex);
							subItemControlData.columnID = ColumnIndexToID(subItemIndex);
							subItemControlData.subItemEditMode = mode;
							subItemControlData.propertyKey = propertyKey;

							if(FAILED(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_ENDSUBITEMEDIT, 0, reinterpret_cast<LPARAM>(&subItemControlData)))) {
								// cancelled by ShellListView, e.g. because the property is read-only
								LVITEMINDEX parentItemIndex = {itemIndex, 0};
								CComPtr<IListViewSubItem> pSubItem = ClassFactory::InitListSubItem(parentItemIndex, subItemIndex, this);

								Raise_CancelSubItemEdit(pSubItem, static_cast<SubItemEditModeConstants>(mode));
							}
						}
					#endif

					PropVariantClear(pPropValue);
				}
			}
		} else {
			LVITEMINDEX parentItemIndex = {itemIndex, 0};
			CComPtr<IListViewSubItem> pSubItem = ClassFactory::InitListSubItem(parentItemIndex, subItemIndex, this);

			Raise_CancelSubItemEdit(pSubItem, static_cast<SubItemEditModeConstants>(mode));
		}

		return pPropertyControl->Destroy();
	#else
		return E_NOTIMPL;
	#endif
}

STDMETHODIMP ExplorerListView::BeginGroupEdit(int /*groupIndex*/, REFIID /*requiredInterface*/, LPVOID* ppObject)
{
	ATLASSUME(ppObject);

	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		/*if(requiredInterface == IID_IPropertyControl) {
			*ppObject = NULL;
			CComObject<DefaultPropertyControl>* pDefaultPropCtlObj = NULL;
			CComObject<DefaultPropertyControl>::CreateInstance(&pDefaultPropCtlObj);
			pDefaultPropCtlObj->AddRef();
			pDefaultPropCtlObj->SetOwner(this);
			pDefaultPropCtlObj->QueryInterface(requiredInterface, ppObject);
			pDefaultPropCtlObj->Release();
			return S_OK;
		}
		ATLASSERT(FALSE && "BeginGroupEdit");*/
		return E_NOINTERFACE;
	#else
		return E_NOTIMPL;
	#endif
}

STDMETHODIMP ExplorerListView::EndGroupEdit(int /*groupIndex*/, int /*mode*/, IPropertyControl* pPropertyControl)
{
	ATLASSUME(pPropertyControl);
	if(!pPropertyControl) {
		return E_POINTER;
	}

	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		return pPropertyControl->Destroy();
	#else
		return E_NOTIMPL;
	#endif
}

STDMETHODIMP ExplorerListView::OnInvokeVerb(int itemIndex, LPCWSTR pVerb)
{
	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		LVITEMINDEX i = {itemIndex, 0};
		CComPtr<IListViewItem> pItem = ClassFactory::InitListItem(i, this, FALSE);
		CComBSTR verb(pVerb);
		Raise_InvokeVerbFromSubItemControl(pItem, verb);
		return S_OK;
	#else
		return E_NOTIMPL;
	#endif
}
// implementation of ISubItemCallback
//////////////////////////////////////////////////////////////////////

#ifdef INCLUDESHELLBROWSERINTERFACE
	//////////////////////////////////////////////////////////////////////
	// implementation of IInternalMessageListener
	HRESULT ExplorerListView::ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam)
	{
		switch(message) {
			case EXLVM_ADDCOLUMN:
				return OnInternalAddColumn(message, wParam, lParam);
				break;
			case EXLVM_COLUMNIDTOINDEX:
				return OnInternalColumnIDToIndex(message, wParam, lParam);
				break;
			case EXLVM_COLUMNINDEXTOID:
				return OnInternalColumnIndexToID(message, wParam, lParam);
				break;
			case EXLVM_GETCOLUMNBYID:
				return OnInternalGetColumnByID(message, wParam, lParam);
				break;
			case EXLVM_SETSORTARROW:
				return OnInternalSetSortArrow(message, wParam, lParam);
				break;
			case EXLVM_GETCOLUMNBITMAP:
				return OnInternalGetColumnBitmap(message, wParam, lParam);
				break;
			case EXLVM_SETCOLUMNBITMAP:
				return OnInternalSetColumnBitmap(message, wParam, lParam);
				break;
			case EXLVM_ADDITEM:
				return OnInternalAddItem(message, wParam, lParam);
				break;
			case EXLVM_ITEMIDTOINDEX:
				return OnInternalItemIDToIndex(message, wParam, lParam);
				break;
			case EXLVM_ITEMINDEXTOID:
				return OnInternalItemIndexToID(message, wParam, lParam);
				break;
			case EXLVM_GETITEMBYID:
				return OnInternalGetItemByID(message, wParam, lParam);
				break;
			case EXLVM_CREATEITEMCONTAINER:
				return OnInternalCreateItemContainer(message, wParam, lParam);
				break;
			case EXLVM_GETITEMIDSFROMVARIANT:
				return OnInternalGetItemIDsFromVariant(message, wParam, lParam);
				break;
			case EXLVM_SETITEMICONINDEX:
				return OnInternalSetItemIconIndex(message, wParam, lParam);
				break;
			case EXLVM_GETITEMPOSITION:
				return OnInternalGetItemPosition(message, wParam, lParam);
				break;
			case EXLVM_SETITEMPOSITION:
				return OnInternalSetItemPosition(message, wParam, lParam);
				break;
			case EXLVM_MOVEDRAGGEDITEMS:
				return OnInternalMoveDraggedItems(message, wParam, lParam);
				break;
			case EXLVM_HITTEST:
				return OnInternalHitTest(message, wParam, lParam);
				break;
			case EXLVM_SORTITEMS:
				return OnInternalSortItems(message, wParam, lParam);
				break;
			case EXLVM_GETSORTORDER:
				return OnInternalGetSortOrder(message, wParam, lParam);
				break;
			case EXLVM_SETSORTORDER:
				return OnInternalSetSortOrder(message, wParam, lParam);
				break;
			case EXLVM_GETCLOSESTINSERTMARKPOSITION:
				return OnInternalGetClosestInsertMarkPosition(message, wParam, lParam);
				break;
			case EXLVM_SETINSERTMARK:
				return OnInternalSetInsertMark(message, wParam, lParam);
				break;
			case EXLVM_SETDROPHILITEDITEM:
				return OnInternalSetDropHilitedItem(message, wParam, lParam);
				break;
			case EXLVM_CONTROLISDRAGSOURCE:
				return OnInternalControlIsDragSource(message, wParam, lParam);
				break;
			case EXLVM_FIREDRAGDROPEVENT:
				return OnInternalFireDragDropEvent(message, wParam, lParam);
				break;
			case EXLVM_OLEDRAG:
				return OnInternalOLEDrag(message, wParam, lParam);
				break;
			case EXLVM_GETIMAGELIST:
				return OnInternalGetImageList(message, wParam, lParam);
				break;
			case EXLVM_SETIMAGELIST:
				return OnInternalSetImageList(message, wParam, lParam);
				break;
			case EXLVM_GETVIEWMODE:
				return OnInternalGetViewMode(message, wParam, lParam);
				break;
			case EXLVM_GETTILEVIEWITEMLINES:
				return OnInternalGetTileViewItemLines(message, wParam, lParam);
				break;
			case EXLVM_GETCOLUMNHEADERVISIBILITY:
				return OnInternalGetColumnHeaderVisibility(message, wParam, lParam);
				break;
			case EXLVM_ISITEMVISIBLE:
				return OnInternalIsItemVisible(message, wParam, lParam);
				break;
		}
		return E_NOTIMPL;
	}
	// implementation of IInternalMessageListener
	//////////////////////////////////////////////////////////////////////
#endif

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ExplorerListView::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_BkImage:
			*pName = GetResString(IDPC_BKIMAGE).Detach();
			return S_OK;
			break;
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
		case PROPCAT_Grouping:
			*pName = GetResString(IDPC_GROUPING).Detach();
			return S_OK;
			break;
		case PROPCAT_HeaderControl:
			*pName = GetResString(IDPC_HEADERCONTROL).Detach();
			return S_OK;
			break;
		case PROPCAT_TileView:
			*pName = GetResString(IDPC_TILEVIEW).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_EXLVW_ALWAYSSHOWSELECTION:
		case DISPID_EXLVW_APPEARANCE:
		case DISPID_EXLVW_BACKGROUNDDRAWMODE:
		case DISPID_EXLVW_BLENDSELECTIONLASSO:
		case DISPID_EXLVW_BORDERSELECT:
		case DISPID_EXLVW_BORDERSTYLE:
		case DISPID_EXLVW_EMPTYMARKUPTEXT:
		case DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT:
		case DISPID_EXLVW_FOOTERINTROTEXT:
		case DISPID_EXLVW_GRIDLINES:
		case DISPID_EXLVW_HIDELABELS:
		case DISPID_EXLVW_HOTMOUSEICON:
		case DISPID_EXLVW_HOTMOUSEPOINTER:
		case DISPID_EXLVW_ITEMHEIGHT:
		case DISPID_EXLVW_MOUSEICON:
		case DISPID_EXLVW_MOUSEPOINTER:
		case DISPID_EXLVW_REGIONAL:
		case DISPID_EXLVW_SHOWSTATEIMAGES:
		case DISPID_EXLVW_SHOWSUBITEMIMAGES:
		case DISPID_EXLVW_VIEW:
		case DISPID_EXLVW_VIRTUALITEMCOUNT:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_EXLVW_ALLOWLABELEDITING:
		case DISPID_EXLVW_ANCHORITEM:
		case DISPID_EXLVW_AUTOARRANGEITEMS:
		case DISPID_EXLVW_AUTOSIZECOLUMNS:
		case DISPID_EXLVW_CALLBACKMASK:
		case DISPID_EXLVW_CARETCOLUMN:
		case DISPID_EXLVW_CARETFOOTERITEM:
		case DISPID_EXLVW_CARETGROUP:
		case DISPID_EXLVW_CARETITEM:
		case DISPID_EXLVW_CHECKITEMONSELECT:
		case DISPID_EXLVW_DISABLEDEVENTS:
		case DISPID_EXLVW_DONTREDRAW:
		case DISPID_EXLVW_DRAWIMAGESASYNCHRONOUSLY:
		case DISPID_EXLVW_EDITHOVERTIME:
		case DISPID_EXLVW_EDITIMEMODE:
		case DISPID_EXLVW_FULLROWSELECT:
		case DISPID_EXLVW_HOTTRACKING:
		case DISPID_EXLVW_HOTTRACKINGHOVERTIME:
		case DISPID_EXLVW_HOVERTIME:
		case DISPID_EXLVW_IMEMODE:
		case DISPID_EXLVW_INCLUDEHEADERINTABORDER:
		case DISPID_EXLVW_INCREMENTALSEARCHSTRING:
		case DISPID_EXLVW_ITEMACTIVATIONMODE:
		case DISPID_EXLVW_ITEMALIGNMENT:
		case DISPID_EXLVW_ITEMBOUNDINGBOXDEFINITION:
		case DISPID_EXLVW_JUSTIFYICONCOLUMNS:
		case DISPID_EXLVW_LABELWRAP:
		case DISPID_EXLVW_MULTISELECT:
		case DISPID_EXLVW_OWNERDRAWN:
		case DISPID_EXLVW_PROCESSCONTEXTMENUKEYS:
		case DISPID_EXLVW_RIGHTTOLEFT:
		case DISPID_EXLVW_SCROLLBARS:
		case DISPID_EXLVW_SELECTEDCOLUMN:
		case DISPID_EXLVW_SIMPLESELECT:
		case DISPID_EXLVW_SINGLEROW:
		case DISPID_EXLVW_SNAPTOGRID:
		case DISPID_EXLVW_SORTORDER:
		case DISPID_EXLVW_TOOLTIPS:
		case DISPID_EXLVW_UNDERLINEDITEMS:
		case DISPID_EXLVW_USEWORKAREAS:
		case DISPID_EXLVW_VIRTUALMODE:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_EXLVW_ABSOLUTEBKIMAGEPOSITION:
		case DISPID_EXLVW_BKIMAGE:
		case DISPID_EXLVW_BKIMAGEPOSITIONX:
		case DISPID_EXLVW_BKIMAGEPOSITIONY:
		case DISPID_EXLVW_BKIMAGESTYLE:
			*pCategory = PROPCAT_BkImage;
			return S_OK;
			break;
		case DISPID_EXLVW_BACKCOLOR:
		case DISPID_EXLVW_EDITBACKCOLOR:
		case DISPID_EXLVW_EDITFORECOLOR:
		case DISPID_EXLVW_FORECOLOR:
		case DISPID_EXLVW_GROUPFOOTERFORECOLOR:
		case DISPID_EXLVW_GROUPHEADERFORECOLOR:
		case DISPID_EXLVW_HOTFORECOLOR:
		case DISPID_EXLVW_INSERTMARKCOLOR:
		case DISPID_EXLVW_OUTLINECOLOR:
		case DISPID_EXLVW_SELECTEDCOLUMNBACKCOLOR:
		case DISPID_EXLVW_TEXTBACKCOLOR:
		case DISPID_EXLVW_TILEVIEWSUBITEMFORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_EXLVW_APPID:
		case DISPID_EXLVW_APPNAME:
		case DISPID_EXLVW_APPSHORTNAME:
		case DISPID_EXLVW_BUILD:
		case DISPID_EXLVW_CHARSET:
		case DISPID_EXLVW_HIMAGELIST:
		case DISPID_EXLVW_HWND:
		case DISPID_EXLVW_HWNDEDIT:
		case DISPID_EXLVW_HWNDHEADER:
		case DISPID_EXLVW_HWNDTOOLTIP:
		case DISPID_EXLVW_ISRELEASE:
		case DISPID_EXLVW_PROGRAMMER:
		case DISPID_EXLVW_TESTER:
		case DISPID_EXLVW_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_EXLVW_ALLOWHEADERDRAGDROP:
		case DISPID_EXLVW_DRAGGEDCOLUMN:
		case DISPID_EXLVW_DRAGGEDITEMS:
		case DISPID_EXLVW_DRAGSCROLLTIMEBASE:
		case DISPID_EXLVW_DROPHILITEDITEM:
		case DISPID_EXLVW_HDRAGIMAGELIST:
		case DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE:
		case DISPID_EXLVW_OLEDRAGIMAGESTYLE:
		case DISPID_EXLVW_REGISTERFOROLEDRAGDROP:
		case DISPID_EXLVW_SHOWDRAGIMAGE:
		case DISPID_EXLVW_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_EXLVW_FONT:
		case DISPID_EXLVW_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_EXLVW_GROUPLOCALE:
		case DISPID_EXLVW_GROUPMARGINBOTTOM:
		case DISPID_EXLVW_GROUPMARGINLEFT:
		case DISPID_EXLVW_GROUPMARGINRIGHT:
		case DISPID_EXLVW_GROUPMARGINTOP:
		case DISPID_EXLVW_GROUPSORTORDER:
		case DISPID_EXLVW_GROUPTEXTPARSINGFLAGS:
		case DISPID_EXLVW_MINITEMROWSVISIBLEINGROUPS:
		case DISPID_EXLVW_SHOWGROUPS:
			*pCategory = PROPCAT_Grouping;
			return S_OK;
			break;
		case DISPID_EXLVW_CLICKABLECOLUMNHEADERS:
		case DISPID_EXLVW_COLUMNHEADERVISIBILITY:
		case DISPID_EXLVW_COLUMNWIDTH:
		case DISPID_EXLVW_FILTERCHANGEDTIMEOUT:
		case DISPID_EXLVW_HEADERFULLDRAGGING:
		case DISPID_EXLVW_HEADERHOTTRACKING:
		case DISPID_EXLVW_HEADERHOVERTIME:
		case DISPID_EXLVW_RESIZABLECOLUMNS:
		case DISPID_EXLVW_SHOWFILTERBAR:
		case DISPID_EXLVW_SHOWHEADERCHEVRON:
		case DISPID_EXLVW_SHOWHEADERSTATEIMAGES:
		case DISPID_EXLVW_USEMINCOLUMNWIDTHS:
			*pCategory = PROPCAT_HeaderControl;
			return S_OK;
			break;
		case DISPID_EXLVW_COLUMNS:
		case DISPID_EXLVW_FIRSTVISIBLEITEM:
		case DISPID_EXLVW_FOOTERITEMS:
		case DISPID_EXLVW_GROUPS:
		case DISPID_EXLVW_HOTITEM:
		case DISPID_EXLVW_LISTITEMS:
		case DISPID_EXLVW_WORKAREAS:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_EXLVW_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
		case DISPID_EXLVW_TILEVIEWITEMLINES:
		case DISPID_EXLVW_TILEVIEWLABELMARGINBOTTOM:
		case DISPID_EXLVW_TILEVIEWLABELMARGINLEFT:
		case DISPID_EXLVW_TILEVIEWLABELMARGINRIGHT:
		case DISPID_EXLVW_TILEVIEWLABELMARGINTOP:
		case DISPID_EXLVW_TILEVIEWTILEHEIGHT:
		case DISPID_EXLVW_TILEVIEWTILEWIDTH:
			*pCategory = PROPCAT_TileView;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString ExplorerListView::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ExplorerListView::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ExplorerListView::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ExplorerListView&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ExplorerListView::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString ExplorerListView::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ExplorerListView::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IGroupComparator
int ExplorerListView::CompareGroups(int groupID1, int groupID2)
{
	int result = 0;

	LVGROUP group1 = {0};
	LVGROUP group2 = {0};
	BOOL earlyAbort = FALSE;
	for(int i = 0; i < 3 && !earlyAbort; ++i) {
		if(sortingSettings.sortingCriteria[i] == sobNone) {
			// there's nothing useful we can do
			earlyAbort = TRUE;
			continue;
		} else if(sortingSettings.sortingCriteria[i] == sobCustom) {
			CComPtr<IListViewGroup> pLvwGroup1 = ClassFactory::InitGroup(groupID1, this, FALSE);
			CComPtr<IListViewGroup> pLvwGroup2 = ClassFactory::InitGroup(groupID2, this, FALSE);
			if(FAILED(Raise_CompareGroups(pLvwGroup1, pLvwGroup2, reinterpret_cast<CompareResultConstants*>(&result)))) {
				result = 0;
				earlyAbort = TRUE;
			} else {
				earlyAbort = (result != 0);
			}
			continue;
		}

		// we'll need some details about the groups
		if((sortingSettings.sortingCriteria[i] != sobShell) && (group1.mask == 0)) {
			group1.cbSize = RunTimeHelper::SizeOf_LVGROUP();
			group1.iGroupId = groupID1;
			group1.mask = LVGF_GROUPID | LVGF_HEADER;
			SendMessage(LVM_GETGROUPINFO, groupID1, reinterpret_cast<LPARAM>(&group1));
			group2.cbSize = RunTimeHelper::SizeOf_LVGROUP();
			group2.iGroupId = groupID2;
			group2.mask = LVGF_GROUPID | LVGF_HEADER;
			SendMessage(LVM_GETGROUPINFO, groupID2, reinterpret_cast<LPARAM>(&group2));
		}
		switch(sortingSettings.sortingCriteria[i]) {
			case sobShell:
				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(shellBrowserInterface.pInternalMessageListener) {
						LONG groups[2] = {groupID1, groupID2};
						BOOL wasHandled = FALSE;
						result = static_cast<short>(HRESULT_CODE(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_COMPAREGROUPS, reinterpret_cast<WPARAM>(groups), reinterpret_cast<LPARAM>(&wasHandled))));
						if(!wasHandled) {
							result = 0;
						}
					}
				#else
					result = 0;
				#endif
				break;
			case sobText:
				if(sortingSettings.localeID != static_cast<LCID>(-1) || sortingSettings.flagsForCompareString != 0) {
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					result = CompareStringW(sortingSettings.localeID, sortingSettings.flagsForCompareString, group1.pszHeader, lstrlenW(group1.pszHeader), group2.pszHeader, lstrlenW(group2.pszHeader)) - 2;
				} else {
					if(sortingSettings.caseSensitive) {
						result = lstrcmpW(group1.pszHeader, group2.pszHeader);
					} else {
						result = lstrcmpiW(group1.pszHeader, group2.pszHeader);
					}
				}
				break;
			case sobNumericIntText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = group1.pszHeader;
					CComBSTR text2 = group2.pszHeader;
					LONG64 int1;
					LONG64 int2;
					if(SUCCEEDED(VarI8FromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int1))) {
						if(SUCCEEDED(VarI8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int2))) {
							if(int1 == int2) {
								result = 0;
							} else {
								result = ((int1 < int2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarI8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareStringW(sortingSettings.localeID, sortingSettings.flagsForCompareString, group1.pszHeader, lstrlenW(group1.pszHeader), group2.pszHeader, lstrlenW(group2.pszHeader)) - 2;
						}
					}
				}
				break;
			case sobNumericFloatText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = group1.pszHeader;
					CComBSTR text2 = group2.pszHeader;
					DOUBLE dbl1;
					DOUBLE dbl2;
					if(SUCCEEDED(VarR8FromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl1))) {
						if(SUCCEEDED(VarR8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl2))) {
							if(dbl1 == dbl2) {
								result = 0;
							} else {
								result = ((dbl1 < dbl2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarR8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareStringW(sortingSettings.localeID, sortingSettings.flagsForCompareString, group1.pszHeader, lstrlenW(group1.pszHeader), group2.pszHeader, lstrlenW(group2.pszHeader)) - 2;
						}
					}
				}
				break;
			case sobCurrencyText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = group1.pszHeader;
					CComBSTR text2 = group2.pszHeader;
					CY curr1;
					CY curr2;
					if(SUCCEEDED(VarCyFromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr1))) {
						if(SUCCEEDED(VarCyFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr2))) {
							if(curr1.int64 == curr2.int64) {
								result = 0;
							} else {
								result = ((curr1.int64 < curr2.int64) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarCyFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareStringW(sortingSettings.localeID, sortingSettings.flagsForCompareString, group1.pszHeader, lstrlenW(group1.pszHeader), group2.pszHeader, lstrlenW(group2.pszHeader)) - 2;
						}
					}
				}
				break;
			case sobDateTimeText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = group1.pszHeader;
					CComBSTR text2 = group2.pszHeader;
					DATE date1;
					DATE date2;
					if(SUCCEEDED(VarDateFromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date1))) {
						if(SUCCEEDED(VarDateFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date2))) {
							if(date1 == date2) {
								result = 0;
							} else {
								result = ((date1 < date2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarDateFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareStringW(sortingSettings.localeID, sortingSettings.flagsForCompareString, group1.pszHeader, lstrlenW(group1.pszHeader), group2.pszHeader, lstrlenW(group2.pszHeader)) - 2;
						}
					}
				}
				break;
		}

		if(result != 0) {
			earlyAbort = TRUE;
		}
	}

	// ensure we don't return illegal values
	if(result < 0) {
		result = -1;
	} else if(result > 0) {
		result = 1;
	}

	if(properties.groupSortOrder == soDescending) {
		result = -result;
	}
	return result;
}
// implementation of IGroupComparator
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IItemComparator
int ExplorerListView::CompareItems(LONG itemID1, LONG itemID2)
{
	// retrieve the items' indexes
	int itemIndex1 = IDToItemIndex(itemID1);
	int itemIndex2 = IDToItemIndex(itemID2);
	if((itemIndex1 == -1) || (itemIndex2 == -1)) {
		// there's nothing useful we can do
		return 0;
	}

	return CompareItemsEx(itemIndex1, itemIndex2);
}

int ExplorerListView::CompareItemsEx(int itemIndex1, int itemIndex2)
{
	int result = 0;

	LVITEM item1 = {0};
	TCHAR pItem1TextBuffer[MAX_ITEMTEXTLENGTH + 1];
	LVITEM item2 = {0};
	TCHAR pItem2TextBuffer[MAX_ITEMTEXTLENGTH + 1];
	BOOL earlyAbort = FALSE;
	for(int i = 0; i < 5 && !earlyAbort; ++i) {
		if(sortingSettings.sortingCriteria[i] == sobNone) {
			// there's nothing useful we can do
			earlyAbort = TRUE;
			continue;
		} else if(sortingSettings.sortingCriteria[i] == sobCustom) {
			LVITEMINDEX i = {itemIndex1, 0};
			CComPtr<IListViewItem> pLvwItem1 = ClassFactory::InitListItem(i, this, FALSE);
			i.iItem = itemIndex2;
			CComPtr<IListViewItem> pLvwItem2 = ClassFactory::InitListItem(i, this, FALSE);
			if(FAILED(Raise_CompareItems(pLvwItem1, pLvwItem2, reinterpret_cast<CompareResultConstants*>(&result)))) {
				result = 0;
				earlyAbort = TRUE;
			} else {
				earlyAbort = (result != 0);
			}
			continue;
		}

		// we'll need some details about the items
		if((sortingSettings.sortingCriteria[i] != sobShell) && (item1.mask == 0)) {
			item1.iItem = itemIndex1;
			if(sortingSettings.column != -1) {
				item1.iSubItem = sortingSettings.column;
			}
			item1.cchTextMax = MAX_ITEMTEXTLENGTH;
			item1.pszText = pItem1TextBuffer;
			item1.stateMask = LVIS_SELECTED | LVIS_STATEIMAGEMASK;
			item1.mask = LVIF_TEXT | LVIF_STATE;
			SendMessage(LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item1));
			item2.iItem = itemIndex2;
			if(sortingSettings.column != -1) {
				item2.iSubItem = sortingSettings.column;
			}
			item2.cchTextMax = MAX_ITEMTEXTLENGTH;
			item2.pszText = pItem2TextBuffer;
			item2.stateMask = LVIS_SELECTED | LVIS_STATEIMAGEMASK;
			item2.mask = LVIF_TEXT | LVIF_STATE;
			SendMessage(LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item2));
		}
		switch(sortingSettings.sortingCriteria[i]) {
			case sobShell:
				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(shellBrowserInterface.pInternalMessageListener) {
						LONG items[2] = {ItemIndexToID(itemIndex1), ItemIndexToID(itemIndex2)};
						BOOL wasHandled = FALSE;
						result = static_cast<short>(HRESULT_CODE(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_COMPAREITEMS, reinterpret_cast<WPARAM>(items), reinterpret_cast<LPARAM>(&wasHandled))));
						if(!wasHandled) {
							result = 0;
						}
					}
				#else
					result = 0;
				#endif
				break;
			case sobText:
				if(sortingSettings.localeID != static_cast<LCID>(-1) || sortingSettings.flagsForCompareString != 0) {
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
				} else {
					if(sortingSettings.caseSensitive) {
						result = lstrcmp(item1.pszText, item2.pszText);
					} else {
						result = lstrcmpi(item1.pszText, item2.pszText);
					}
				}
				break;
			case sobNumericIntText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					LONG64 int1;
					LONG64 int2;
					if(SUCCEEDED(VarI8FromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int1))) {
						if(SUCCEEDED(VarI8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int2))) {
							if(int1 == int2) {
								result = 0;
							} else {
								result = ((int1 < int2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarI8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &int2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobNumericFloatText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					DOUBLE dbl1;
					DOUBLE dbl2;
					if(SUCCEEDED(VarR8FromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl1))) {
						if(SUCCEEDED(VarR8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl2))) {
							if(dbl1 == dbl2) {
								result = 0;
							} else {
								result = ((dbl1 < dbl2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarR8FromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &dbl2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobDateTimeText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					DATE date1;
					DATE date2;
					if(SUCCEEDED(VarDateFromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date1))) {
						if(SUCCEEDED(VarDateFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date2))) {
							if(date1 == date2) {
								result = 0;
							} else {
								result = ((date1 < date2) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarDateFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &date2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobCurrencyText:
				{
					if(sortingSettings.localeID == static_cast<LCID>(-1)) {
						sortingSettings.localeID = GetThreadLocale();
					}
					CComBSTR text1 = item1.pszText;
					CComBSTR text2 = item2.pszText;
					CY curr1;
					CY curr2;
					if(SUCCEEDED(VarCyFromStr(text1, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr1))) {
						if(SUCCEEDED(VarCyFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr2))) {
							if(curr1.int64 == curr2.int64) {
								result = 0;
							} else {
								result = ((curr1.int64 < curr2.int64) ? -1 : 1);
							}
						} else {
							// text2 is illegal - move it to the end
							result = -1;
						}
					} else {
						// text1 is illegal
						if(SUCCEEDED(VarCyFromStr(text2, sortingSettings.localeID, sortingSettings.flagsForVarFromStr, &curr2))) {
							// text2 is valid - move it to the front
							result = 1;
						} else {
							// both are illegal - sort by text
							result = CompareString(sortingSettings.localeID, sortingSettings.flagsForCompareString, item1.pszText, lstrlen(item1.pszText), item2.pszText, lstrlen(item2.pszText)) - 2;
						}
					}
				}
				break;
			case sobSelectionState:
				if((item1.state & LVIS_SELECTED) == (item2.state & LVIS_SELECTED)) {
					result = 0;
				} else if(item1.state & LVIS_SELECTED) {
					result = -1;
				} else {
					result = 1;
				}
				break;
			case sobStateImageIndex:
				if(((item1.state & LVIS_STATEIMAGEMASK) >> 12) == ((item2.state & LVIS_STATEIMAGEMASK) >> 12)) {
					result = 0;
				} else {
					result = (((item1.state & LVIS_STATEIMAGEMASK) >> 12) > ((item2.state & LVIS_STATEIMAGEMASK) >> 12) ? -1 : 1);
				}
				break;
		}

		if(result != 0) {
			earlyAbort = TRUE;
		}
	}

	// ensure we don't return illegal values
	if(result < 0) {
		result = -1;
	} else if(result > 0) {
		result = 1;
	}

	if(properties.sortOrder == soDescending) {
		result = -result;
	}
	return result;
}
// implementation of IItemComparator
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP ExplorerListView::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_EXLVW_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_EXLVW_AUTOARRANGEITEMS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.autoArrangeItems), description);
			break;
		case DISPID_EXLVW_BACKGROUNDDRAWMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.backgroundDrawMode), description);
			break;
		case DISPID_EXLVW_BKIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.bkImageStyle), description);
			break;
		case DISPID_EXLVW_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_EXLVW_COLUMNHEADERVISIBILITY:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.columnHeaderVisibility), description);
			break;
		case DISPID_EXLVW_EDITIMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.editIMEMode), description);
			break;
		case DISPID_EXLVW_EMPTYMARKUPTEXT:
		case DISPID_EXLVW_FOOTERINTROTEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.emptyMarkupTextAlignment), description);
			break;
		case DISPID_EXLVW_FULLROWSELECT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.fullRowSelect), description);
			break;
		case DISPID_EXLVW_GROUPSORTORDER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.groupSortOrder), description);
			break;
		case DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.headerOLEDragImageStyle), description);
			break;
		case DISPID_EXLVW_HOTFORECOLOR:
			if(properties.hotForeColor == -1) {
				hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hotForeColor), description);
			} else {
				return IPerPropertyBrowsingImpl<ExplorerListView>::GetDisplayString(property, pDescription);
			}
			break;
		case DISPID_EXLVW_HOTMOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hotMousePointer), description);
			break;
		case DISPID_EXLVW_IMEMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.IMEMode), description);
			break;
		case DISPID_EXLVW_ITEMACTIVATIONMODE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.itemActivationMode), description);
			break;
		case DISPID_EXLVW_ITEMALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.itemAlignment), description);
			break;
		case DISPID_EXLVW_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_EXLVW_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_EXLVW_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_EXLVW_SCROLLBARS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.scrollBars), description);
			break;
		case DISPID_EXLVW_SELECTEDCOLUMNBACKCOLOR:
			if(properties.selectedColumnBackColor == -1) {
				hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.selectedColumnBackColor), description);
			} else {
				return IPerPropertyBrowsingImpl<ExplorerListView>::GetDisplayString(property, pDescription);
			}
			break;
		case DISPID_EXLVW_SORTORDER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.sortOrder), description);
			break;
		case DISPID_EXLVW_TEXTBACKCOLOR:
			if(properties.textBackColor == CLR_NONE) {
				hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.textBackColor), description);
			} else {
				return IPerPropertyBrowsingImpl<ExplorerListView>::GetDisplayString(property, pDescription);
			}
			break;
		case DISPID_EXLVW_TILEVIEWSUBITEMFORECOLOR:
			if(properties.tileViewSubItemForeColor == -1) {
				hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.tileViewSubItemForeColor), description);
			} else {
				return IPerPropertyBrowsingImpl<ExplorerListView>::GetDisplayString(property, pDescription);
			}
			break;
		case DISPID_EXLVW_TOOLTIPS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.toolTips), description);
			break;
		case DISPID_EXLVW_UNDERLINEDITEMS:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.underlinedItems), description);
			break;
		case DISPID_EXLVW_VIEW:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.view), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ExplorerListView>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ExplorerListView::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_EXLVW_BORDERSTYLE:
		case DISPID_EXLVW_GROUPSORTORDER:
		case DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE:
		#ifndef ALLOWBOTTOMRIGHTALIGNMENT
			case DISPID_EXLVW_ITEMALIGNMENT:
		#endif
		case DISPID_EXLVW_OLEDRAGIMAGESTYLE:
		case DISPID_EXLVW_SORTORDER:
			c = 2;
			break;
		case DISPID_EXLVW_APPEARANCE:
		case DISPID_EXLVW_AUTOARRANGEITEMS:
		case DISPID_EXLVW_BACKGROUNDDRAWMODE:
		case DISPID_EXLVW_BKIMAGESTYLE:
		case DISPID_EXLVW_COLUMNHEADERVISIBILITY:
		case DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT:
		case DISPID_EXLVW_FULLROWSELECT:
		case DISPID_EXLVW_ITEMACTIVATIONMODE:
		case DISPID_EXLVW_SCROLLBARS:
			c = 3;
			break;
		#ifdef ALLOWBOTTOMRIGHTALIGNMENT
			case DISPID_EXLVW_ITEMALIGNMENT:
		#endif
		case DISPID_EXLVW_RIGHTTOLEFT:
		case DISPID_EXLVW_TOOLTIPS:
		case DISPID_EXLVW_UNDERLINEDITEMS:
			c = 4;
			break;
		case DISPID_EXLVW_VIEW:
			c = 6;
			break;
		case DISPID_EXLVW_EDITIMEMODE:
		case DISPID_EXLVW_IMEMODE:
			c = 12;
			break;
		case DISPID_EXLVW_HOTMOUSEPOINTER:
		case DISPID_EXLVW_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ExplorerListView>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = reinterpret_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = reinterpret_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if(((property == DISPID_EXLVW_HOTMOUSEPOINTER) || (property == DISPID_EXLVW_MOUSEPOINTER)) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if((property == DISPID_EXLVW_IMEMODE) || (property == DISPID_EXLVW_EDITIMEMODE)) {
			// the enum is -1-based
			--propertyValue;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = reinterpret_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// just use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_EXLVW_APPEARANCE:
		case DISPID_EXLVW_AUTOARRANGEITEMS:
		case DISPID_EXLVW_BACKGROUNDDRAWMODE:
		case DISPID_EXLVW_BKIMAGESTYLE:
		case DISPID_EXLVW_BORDERSTYLE:
		case DISPID_EXLVW_COLUMNHEADERVISIBILITY:
		case DISPID_EXLVW_EDITIMEMODE:
		case DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT:
		case DISPID_EXLVW_FULLROWSELECT:
		case DISPID_EXLVW_GROUPSORTORDER:
		case DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE:
		case DISPID_EXLVW_HOTMOUSEPOINTER:
		case DISPID_EXLVW_IMEMODE:
		case DISPID_EXLVW_ITEMACTIVATIONMODE:
		case DISPID_EXLVW_ITEMALIGNMENT:
		case DISPID_EXLVW_MOUSEPOINTER:
		case DISPID_EXLVW_OLEDRAGIMAGESTYLE:
		case DISPID_EXLVW_RIGHTTOLEFT:
		case DISPID_EXLVW_SCROLLBARS:
		case DISPID_EXLVW_SORTORDER:
		case DISPID_EXLVW_TOOLTIPS:
		case DISPID_EXLVW_UNDERLINEDITEMS:
		case DISPID_EXLVW_VIEW:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		case DISPID_EXLVW_HOTFORECOLOR:
		case DISPID_EXLVW_SELECTEDCOLUMNBACKCOLOR:
		case DISPID_EXLVW_TILEVIEWSUBITEMFORECOLOR:
			if(cookie == -1) {
				VariantInit(pPropertyValue);
				pPropertyValue->vt = VT_I4;
				// we used the property value itself as cookie
				pPropertyValue->lVal = cookie;
			} else {
				return IPerPropertyBrowsingImpl<ExplorerListView>::GetPredefinedValue(property, cookie, pPropertyValue);
			}
			break;
		case DISPID_EXLVW_TEXTBACKCOLOR:
			if(cookie == CLR_NONE) {
				VariantInit(pPropertyValue);
				pPropertyValue->vt = VT_I4;
				// we used the property value itself as cookie
				pPropertyValue->lVal = cookie;
			} else {
				return IPerPropertyBrowsingImpl<ExplorerListView>::GetPredefinedValue(property, cookie, pPropertyValue);
			}
			break;
		default:
			return IPerPropertyBrowsingImpl<ExplorerListView>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_EXLVW_EMPTYMARKUPTEXT:
		case DISPID_EXLVW_FOOTERINTROTEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<ExplorerListView>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ExplorerListView::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_EXLVW_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_AUTOARRANGEITEMS:
			switch(cookie) {
				case aaiNone:
					description = GetResStringWithNumber(IDP_AUTOARRANGEITEMSNONE, aaiNone);
					break;
				case aaiNormal:
					description = GetResStringWithNumber(IDP_AUTOARRANGEITEMSNORMAL, aaiNormal);
					break;
				case aaiIntelligent:
					description = GetResStringWithNumber(IDP_AUTOARRANGEITEMSINTELLIGENT, aaiIntelligent);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_BACKGROUNDDRAWMODE:
			switch(cookie) {
				case bdmNormal:
					description = GetResStringWithNumber(IDP_BACKGROUNDDRAWMODENORMAL, bdmNormal);
					break;
				case bdmByParent:
					description = GetResStringWithNumber(IDP_BACKGROUNDDRAWMODEBYPARENT, bdmByParent);
					break;
				case bdmByParentWithShadowText:
					description = GetResStringWithNumber(IDP_BACKGROUNDDRAWMODEBYPARENTWITHSHADOWTEXT, bdmByParentWithShadowText);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_BKIMAGESTYLE:
			switch(cookie) {
				case bisNormal:
					description = GetResStringWithNumber(IDP_BKIMAGESTYLENORMAL, bisNormal);
					break;
				case bisTiled:
					description = GetResStringWithNumber(IDP_BKIMAGESTYLETILED, bisTiled);
					break;
				case bisWatermark:
					description = GetResStringWithNumber(IDP_BKIMAGESTYLEWATERMARK, bisWatermark);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_COLUMNHEADERVISIBILITY:
			switch(cookie) {
				case chvInvisible:
					description = GetResStringWithNumber(IDP_COLUMNHEADERVISIBILITYINVISIBLE, chvInvisible);
					break;
				case chvVisibleInDetailsView:
					description = GetResStringWithNumber(IDP_COLUMNHEADERVISIBILITYVISIBLEINDETAILSVIEW, chvVisibleInDetailsView);
					break;
				case chvVisibleInAllViews:
					description = GetResStringWithNumber(IDP_COLUMNHEADERVISIBILITYVISIBLEINALLVIEWS, chvVisibleInAllViews);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_EDITIMEMODE:
		case DISPID_EXLVW_IMEMODE:
			switch(cookie) {
				case imeInherit:
					description = GetResStringWithNumber(IDP_IMEMODEINHERIT, imeInherit);
					break;
				case imeNoControl:
					description = GetResStringWithNumber(IDP_IMEMODENOCONTROL, imeNoControl);
					break;
				case imeOn:
					description = GetResStringWithNumber(IDP_IMEMODEON, imeOn);
					break;
				case imeOff:
					description = GetResStringWithNumber(IDP_IMEMODEOFF, imeOff);
					break;
				case imeDisable:
					description = GetResStringWithNumber(IDP_IMEMODEDISABLE, imeDisable);
					break;
				case imeHiragana:
					description = GetResStringWithNumber(IDP_IMEMODEHIRAGANA, imeHiragana);
					break;
				case imeKatakana:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANA, imeKatakana);
					break;
				case imeKatakanaHalf:
					description = GetResStringWithNumber(IDP_IMEMODEKATAKANAHALF, imeKatakanaHalf);
					break;
				case imeAlphaFull:
					description = GetResStringWithNumber(IDP_IMEMODEALPHAFULL, imeAlphaFull);
					break;
				case imeAlpha:
					description = GetResStringWithNumber(IDP_IMEMODEALPHA, imeAlpha);
					break;
				case imeHangulFull:
					description = GetResStringWithNumber(IDP_IMEMODEHANGULFULL, imeHangulFull);
					break;
				case imeHangul:
					description = GetResStringWithNumber(IDP_IMEMODEHANGUL, imeHangul);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT:
			switch(cookie) {
				case alLeft:
					description = GetResStringWithNumber(IDP_ALIGNMENTLEFT, alLeft);
					break;
				case alCenter:
					description = GetResStringWithNumber(IDP_ALIGNMENTCENTER, alCenter);
					break;
				case alRight:
					description = GetResStringWithNumber(IDP_ALIGNMENTRIGHT, alRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_FULLROWSELECT:
			switch(cookie) {
				case frsDisabled:
					description = GetResStringWithNumber(IDP_FULLROWSELECTDISABLED, frsDisabled);
					break;
				case frsNormalMode:
					description = GetResStringWithNumber(IDP_FULLROWSELECTNORMALMODE, frsNormalMode);
					break;
				case frsExtendedMode:
					description = GetResStringWithNumber(IDP_FULLROWSELECTEXTENDEDMODE, frsExtendedMode);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_GROUPSORTORDER:
		case DISPID_EXLVW_SORTORDER:
			switch(cookie) {
				case soAscending:
					description = GetResStringWithNumber(IDP_SORTORDERASCENDING, soAscending);
					break;
				case soDescending:
					description = GetResStringWithNumber(IDP_SORTORDERDESCENDING, soDescending);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE:
		case DISPID_EXLVW_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_HOTFORECOLOR:
		case DISPID_EXLVW_SELECTEDCOLUMNBACKCOLOR:
		case DISPID_EXLVW_TILEVIEWSUBITEMFORECOLOR:
			switch(cookie) {
				case -1:
					description = GetResString(IDP_SYSTEMDEFAULTCOLOR);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_HOTMOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_ITEMACTIVATIONMODE:
			switch(cookie) {
				case iamOneSingleClick:
					description = GetResStringWithNumber(IDP_ITEMACTIVATIONMODEONESINGLECLICK, iamOneSingleClick);
					break;
				case iamTwoSingleClicks:
					description = GetResStringWithNumber(IDP_ITEMACTIVATIONMODETWOSINGLECLICKS, iamTwoSingleClicks);
					break;
				case iamOneDoubleClick:
					description = GetResStringWithNumber(IDP_ITEMACTIVATIONMODEONEDOUBLECLICK, iamOneDoubleClick);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_ITEMALIGNMENT:
			switch(cookie) {
				case iaTop:
					description = GetResStringWithNumber(IDP_ITEMALIGNMENTTOP, iaTop);
					break;
				case iaLeft:
					description = GetResStringWithNumber(IDP_ITEMALIGNMENTLEFT, iaLeft);
					break;
				#ifdef ALLOWBOTTOMRIGHTALIGNMENT
					case iaBottom:
						description = GetResStringWithNumber(IDP_ITEMALIGNMENTBOTTOM, iaBottom);
						break;
					case iaRight:
						description = GetResStringWithNumber(IDP_ITEMALIGNMENTRIGHT, iaRight);
						break;
				#endif
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_SCROLLBARS:
			switch(cookie) {
				case sbNone:
					description = GetResStringWithNumber(IDP_SCROLLBARSNONE, sbNone);
					break;
				case sbNormal:
					description = GetResStringWithNumber(IDP_SCROLLBARSNORMAL, sbNormal);
					break;
				case sbFlat:
					description = GetResStringWithNumber(IDP_SCROLLBARSFLAT, sbFlat);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_TEXTBACKCOLOR:
			switch(cookie) {
				case CLR_NONE:
					description = GetResString(IDP_TRANSPARENTCOLOR);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_TOOLTIPS:
			switch(cookie) {
				case ttLabelTipsIconsAndTilesOnly:
					description = GetResStringWithNumber(IDP_TOOLTIPSLABELTIPSSOME, ttLabelTipsIconsAndTilesOnly);
					break;
				case ttLabelTipsAlways:
					description = GetResStringWithNumber(IDP_TOOLTIPSLABELTIPSALWAYS, ttLabelTipsAlways);
					break;
				case ttInfoTips:
					description = GetResStringWithNumber(IDP_TOOLTIPSINFOTIPS, ttInfoTips);
					break;
				case ttLabelTipsAlways | ttInfoTips:
					description = GetResStringWithNumber(IDP_TOOLTIPSLABELANDINFOTIPS, ttLabelTipsAlways | ttInfoTips);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_UNDERLINEDITEMS:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_UNDERLINEDITEMSNONE, 0);
					break;
				case uiHot:
					description = GetResStringWithNumber(IDP_UNDERLINEDITEMSHOT, uiHot);
					break;
				case uiCold:
					description = GetResStringWithNumber(IDP_UNDERLINEDITEMSCOLD, uiCold);
					break;
				case uiHot | uiCold:
					description = GetResStringWithNumber(IDP_UNDERLINEDITEMSHOTCOLD, uiHot | uiCold);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_EXLVW_VIEW:
			switch(cookie) {
				case vIcons:
					description = GetResStringWithNumber(IDP_VIEWICONS, vIcons);
					break;
				case vSmallIcons:
					description = GetResStringWithNumber(IDP_VIEWSMALLICONS, vSmallIcons);
					break;
				case vList:
					description = GetResStringWithNumber(IDP_VIEWLIST, vList);
					break;
				case vDetails:
					description = GetResStringWithNumber(IDP_VIEWDETAILS, vDetails);
					break;
				case vTiles:
					description = GetResStringWithNumber(IDP_VIEWTILES, vTiles);
					break;
				case vExtendedTiles:
					description = GetResStringWithNumber(IDP_VIEWEXTENDEDTILES, vExtendedTiles);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP ExplorerListView::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = reinterpret_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ExplorerListView::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ExplorerListView>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ExplorerListView::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP ExplorerListView::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT ExplorerListView::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT ExplorerListView::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_EXLVW_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ExplorerListView::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ExplorerListView::get_AbsoluteBkImagePosition(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.absoluteBkImagePosition);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_AbsoluteBkImagePosition(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ABSOLUTEBKIMAGEPOSITION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.absoluteBkImagePosition != b) {
		properties.absoluteBkImagePosition = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			LVBKIMAGE bkImage = {0};
			bkImage.cchImageMax = MAX_PATH;
			TCHAR pBuffer[MAX_PATH + 1];
			bkImage.pszImage = pBuffer;
			SendMessage(LVM_GETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
			if(properties.bkImageStyle == bisWatermark) {
				/* Windows' watermark style implementation is crappy - a listview that has the
				   LVBKIF_TYPE_WATERMARK flag set, won't handle LVM_GETBKIMAGE. */
				bkImage.cchImageMax = 0;
				bkImage.hbm = flags.hCurrentBackgroundBitmap;
				bkImage.xOffsetPercent = properties.bkImagePositionX;
				bkImage.yOffsetPercent = properties.bkImagePositionY;
				bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
			}

			BOOL b = FALSE;
			if(bkImage.hbm) {
				bkImage.cchImageMax = 0;
				/* SysListView32 destroys its current bitmap if it receives LVM_SETBKIMAGE, so re-assign a copy
				   of the bitmap. */
				HBITMAP hBufferedBitmap = bkImage.hbm;
				bkImage.hbm = CopyBitmap(bkImage.hbm);

				VARIANT v;
				VariantInit(&v);
				if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4)) && hBufferedBitmap == static_cast<HBITMAP>(LongToHandle(v.lVal))) {
					// update property
					VariantClear(&properties.bkImage);
					properties.bkImage.lVal = HandleToLong(bkImage.hbm);
					properties.bkImage.vt = VT_I4;
					b = TRUE;
				}
			}

			if(properties.absoluteBkImagePosition) {
				bkImage.ulFlags |= LVBKIF_FLAG_TILEOFFSET;
			} else {
				bkImage.ulFlags &= ~LVBKIF_FLAG_TILEOFFSET;
			}

			// NOTE: We had wParam set to TRUE here. Why?
			SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
			properties.ownsBkImageBitmap = b;
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_ABSOLUTEBKIMAGEPOSITION);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AllowHeaderDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowHeaderDragDrop = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_HEADERDRAGDROP) == LVS_EX_HEADERDRAGDROP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowHeaderDragDrop);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_AllowHeaderDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ALLOWHEADERDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowHeaderDragDrop != b) {
		properties.allowHeaderDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowHeaderDragDrop) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERDRAGDROP, LVS_EX_HEADERDRAGDROP);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERDRAGDROP, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_ALLOWHEADERDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AllowLabelEditing(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowLabelEditing = ((GetStyle() & LVS_EDITLABELS) == LVS_EDITLABELS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowLabelEditing);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_AllowLabelEditing(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ALLOWLABELEDITING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowLabelEditing != b) {
		properties.allowLabelEditing = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowLabelEditing) {
				ModifyStyle(0, LVS_EDITLABELS);
			} else {
				ModifyStyle(LVS_EDITLABELS, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_ALLOWLABELEDITING);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AlwaysShowSelection(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysShowSelection = ((GetStyle() & LVS_SHOWSELALWAYS) == LVS_SHOWSELALWAYS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysShowSelection);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_AlwaysShowSelection(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ALWAYSSHOWSELECTION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysShowSelection != b) {
		properties.alwaysShowSelection = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.alwaysShowSelection) {
				ModifyStyle(0, LVS_SHOWSELALWAYS, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				ModifyStyle(LVS_SHOWSELALWAYS, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_ALWAYSSHOWSELECTION);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AnchorItem(IListViewItem** ppAnchorItem)
{
	ATLASSERT_POINTER(ppAnchorItem, IListViewItem*);
	if(!ppAnchorItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HRESULT hr = E_FAIL;
		LVITEMINDEX anchorItem = {-1, 0};
		CComPtr<IListView_WIN7> pListView7 = NULL;
		CComPtr<IListView_WINVISTA> pListViewVista = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
			ATLASSUME(pListView7);
			hr = pListView7->GetSelectionMark(&anchorItem);
		} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
			ATLASSUME(pListViewVista);
			hr = pListViewVista->GetSelectionMark(&anchorItem);
		} else {
			anchorItem.iItem = static_cast<int>(SendMessage(LVM_GETSELECTIONMARK, 0, 0));
			hr = S_OK;
		}
		if(SUCCEEDED(hr)) {
			ClassFactory::InitListItem(anchorItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppAnchorItem));
		}
		return hr;
	}

	return E_FAIL;
}

STDMETHODIMP ExplorerListView::putref_AnchorItem(IListViewItem* pNewAnchorItem)
{
	PUTPROPPROLOG(DISPID_EXLVW_ANCHORITEM);
	int newAnchorItem = -1;
	if(pNewAnchorItem) {
		LONG l = -1;
		pNewAnchorItem->get_Index(&l);
		newAnchorItem = l;
		// TODO: Shouldn't we AddRef' pNewAnchorItem?
	}

	if(IsWindow()) {
		SendMessage(LVM_SETSELECTIONMARK, 0, newAnchorItem);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_ANCHORITEM);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_APPEARANCE);
	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 3;
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ExplorerListView");
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ExLVw");
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AutoArrangeItems(AutoArrangeItemsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoArrangeItemsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.autoArrangeItems = aaiNone;
		if(GetStyle() & LVS_AUTOARRANGE) {
			properties.autoArrangeItems = aaiNormal;
		} else if(static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)) & LVS_EX_AUTOAUTOARRANGE) {
			properties.autoArrangeItems = aaiIntelligent;
		}
	}

	*pValue = properties.autoArrangeItems;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_AutoArrangeItems(AutoArrangeItemsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_AUTOARRANGEITEMS);
	if(properties.autoArrangeItems != newValue) {
		properties.autoArrangeItems = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.autoArrangeItems) {
				case aaiNone:
					ModifyStyle(LVS_AUTOARRANGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOAUTOARRANGE, 0);
					break;
				case aaiNormal:
					ModifyStyle(0, LVS_AUTOARRANGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOAUTOARRANGE, 0);
					break;
				case aaiIntelligent:
					if(IsComctl32Version610OrNewer()) {
						ModifyStyle(LVS_AUTOARRANGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOAUTOARRANGE, LVS_EX_AUTOAUTOARRANGE);
					} else {
						ModifyStyle(0, LVS_AUTOARRANGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOAUTOARRANGE, 0);
					}
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_AUTOARRANGEITEMS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_AutoSizeColumns(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.autoSizeColumns = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_AUTOSIZECOLUMNS) == LVS_EX_AUTOSIZECOLUMNS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoSizeColumns);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_AutoSizeColumns(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_AUTOSIZECOLUMNS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoSizeColumns != b) {
		properties.autoSizeColumns = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.autoSizeColumns) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOSIZECOLUMNS, LVS_EX_AUTOSIZECOLUMNS);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOSIZECOLUMNS, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_AUTOSIZECOLUMNS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(LVM_GETBKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.backColor = 0x80000000 | COLOR_WINDOW;
		} else if(color != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = color;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BackgroundDrawMode(BackgroundDrawModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BackgroundDrawModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		DWORD style = static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0));
		if(style & (LVS_EX_TRANSPARENTBKGND | LVS_EX_TRANSPARENTSHADOWTEXT)) {
			properties.backgroundDrawMode = bdmByParentWithShadowText;
		} else if(style & LVS_EX_TRANSPARENTBKGND) {
			properties.backgroundDrawMode = bdmByParent;
		} else {
			properties.backgroundDrawMode = bdmNormal;
		}
	}

	*pValue = properties.backgroundDrawMode;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BackgroundDrawMode(BackgroundDrawModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BACKGROUNDDRAWMODE);
	if(properties.backgroundDrawMode != newValue) {
		properties.backgroundDrawMode = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			switch(properties.backgroundDrawMode) {
				case bdmNormal:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, (LVS_EX_TRANSPARENTBKGND | LVS_EX_TRANSPARENTSHADOWTEXT), 0);
					break;
				case bdmByParent:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, (LVS_EX_TRANSPARENTBKGND | LVS_EX_TRANSPARENTSHADOWTEXT), LVS_EX_TRANSPARENTBKGND);
					break;
				case bdmByParentWithShadowText:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, (LVS_EX_TRANSPARENTBKGND | LVS_EX_TRANSPARENTSHADOWTEXT), (LVS_EX_TRANSPARENTBKGND | LVS_EX_TRANSPARENTSHADOWTEXT));
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_BACKGROUNDDRAWMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BkImage(VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	VariantClear(pValue);
	VariantCopy(pValue, &properties.bkImage);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BkImage(VARIANT newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BKIMAGE);
	HBITMAP hPreviousBitmap = NULL;
	VARIANT v;
	VariantInit(&v);
	if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4))) {
		hPreviousBitmap = static_cast<HBITMAP>(LongToHandle(v.lVal));
	}

	switch(newValue.vt) {
		case VT_BSTR:
			if(IsWindow()) {
				LVBKIMAGE bkImage = {0};

				BSTR tmp = NULL;
				if(newValue.vt == VT_BSTR) {
					tmp = newValue.bstrVal;
				}
				COLE2T converter(tmp);
				bkImage.pszImage = converter;
				bkImage.cchImageMax = lstrlen(bkImage.pszImage);

				if(bkImage.cchImageMax) {
					bkImage.ulFlags = LVBKIF_SOURCE_URL;
					if(properties.absoluteBkImagePosition) {
						bkImage.ulFlags |= LVBKIF_FLAG_TILEOFFSET;
					}
					switch(properties.bkImageStyle) {
						case bisWatermark:
							// not supported - fall through
						case bisNormal:
							bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
							break;
						case bisTiled:
							bkImage.ulFlags |= LVBKIF_STYLE_TILE;
							break;
					}
					bkImage.xOffsetPercent = properties.bkImagePositionX;
					bkImage.yOffsetPercent = properties.bkImagePositionY;
				} else {
					bkImage.pszImage = NULL;
					switch(properties.bkImageStyle) {
						case bisWatermark:
							if(RunTimeHelper::IsCommCtrl6()) {
								bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
								break;
							}
							// not supported - fall through
						case bisNormal:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL;
							break;
						case bisTiled:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_TILE;
							break;
					}
				}
				// OnListSetBkImage will synchronize properties.bkImage
				SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
				// properties.bkImage now contains a bitmap handle
			} else if(!IsInDesignMode()) {
				if(properties.ownsBkImageBitmap && hPreviousBitmap && GetObjectType(hPreviousBitmap) == OBJ_BITMAP) {
					DeleteObject(hPreviousBitmap);
				}
				VariantClear(&properties.bkImage);
				VariantCopy(&properties.bkImage, &newValue);
				// TODO: Is this correct?
				properties.ownsBkImageBitmap = TRUE;
			}
			break;
		case VT_DISPATCH: {
			HBITMAP hNewBitmap = NULL;
			if(newValue.pdispVal) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(newValue.pdispVal);
				if(pPicture) {
					SHORT pictureType = PICTYPE_NONE;
					pPicture->get_Type(&pictureType);
					if(pictureType == PICTYPE_BITMAP) {
						OLE_HANDLE h = NULL;
						pPicture->get_Handle(&h);

						hNewBitmap = CopyBitmap(static_cast<HBITMAP>(LongToHandle(h)));
					}
				}
			}

			if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
				LVBKIMAGE bkImage = {0};

				if(hNewBitmap) {
					bkImage.ulFlags = LVBKIF_SOURCE_HBITMAP;
					if(properties.absoluteBkImagePosition) {
						bkImage.ulFlags |= LVBKIF_FLAG_TILEOFFSET;
					}
					switch(properties.bkImageStyle) {
						case bisNormal:
							bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
							break;
						case bisTiled:
							bkImage.ulFlags |= LVBKIF_STYLE_TILE;
							break;
						case bisWatermark:
							// yes, we've to set LVBKIF_SOURCE_NONE - the MS employee must have been stoned...
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
							break;
					}
					bkImage.hbm = hNewBitmap;
					bkImage.xOffsetPercent = properties.bkImagePositionX;
					bkImage.yOffsetPercent = properties.bkImagePositionY;
				} else {
					switch(properties.bkImageStyle) {
						case bisWatermark:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
							break;
						case bisNormal:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL;
							break;
						case bisTiled:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_TILE;
							break;
					}
				}
				// OnListSetBkImage will synchronize properties.bkImage
				SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
				// properties.bkImage now contains a bitmap handle
			} else {
				if(properties.ownsBkImageBitmap && hPreviousBitmap && GetObjectType(hPreviousBitmap) == OBJ_BITMAP) {
					DeleteObject(hPreviousBitmap);
				}
				VariantClear(&properties.bkImage);
			}

			// create a picture object out of hNewBitmap
			if(hNewBitmap) {
				PICTDESC picture = {0};
				picture.cbSizeofstruct = sizeof(picture);
				picture.bmp.hbitmap = hNewBitmap;
				picture.picType = PICTYPE_BITMAP;

				// now create an IDispatch object out of the bitmap
				OleCreatePictureIndirect(&picture, IID_IDispatch, TRUE, reinterpret_cast<LPVOID*>(&properties.bkImage.pdispVal));
				properties.ownsBkImageBitmap = TRUE;
			} else {
				properties.bkImage.pdispVal = NULL;
				properties.ownsBkImageBitmap = FALSE;
			}
			properties.bkImage.vt = VT_DISPATCH;
			break;
		}
		case VT_I2:
		case VT_UI2:
		case VT_I4:
		case VT_UI4:
		case VT_INT:
		case VT_UINT: {
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &newValue, 0, VT_I4))) {
				HBITMAP hNewBitmap = CopyBitmap(static_cast<HBITMAP>(LongToHandle(newValue.lVal)));
				if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
					LVBKIMAGE bkImage = {0};

					if(hNewBitmap) {
						bkImage.ulFlags = LVBKIF_SOURCE_HBITMAP;
						if(properties.absoluteBkImagePosition) {
							bkImage.ulFlags |= LVBKIF_FLAG_TILEOFFSET;
						}
						switch(properties.bkImageStyle) {
							case bisNormal:
								bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
								break;
							case bisTiled:
								bkImage.ulFlags |= LVBKIF_STYLE_TILE;
								break;
							case bisWatermark:
								// yes, we've to set LVBKIF_SOURCE_NONE - the MS employee must have been stoned...
								bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
								break;
						}
						bkImage.hbm = hNewBitmap;
						bkImage.xOffsetPercent = properties.bkImagePositionX;
						bkImage.yOffsetPercent = properties.bkImagePositionY;
					} else {
						switch(properties.bkImageStyle) {
							case bisWatermark:
								bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
								break;
							case bisNormal:
								bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL;
								break;
							case bisTiled:
								bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_TILE;
								break;
						}
					}
					// OnListSetBkImage will synchronize properties.bkImage
					SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
					properties.ownsBkImageBitmap = TRUE;
					// properties.bkImage now contains a bitmap handle
				} else {
					if(properties.ownsBkImageBitmap && hPreviousBitmap && GetObjectType(hPreviousBitmap) == OBJ_BITMAP) {
						DeleteObject(hPreviousBitmap);
					}
					VariantClear(&properties.bkImage);
					properties.bkImage.lVal = HandleToLong(hNewBitmap);
					properties.bkImage.vt = VT_I4;
					properties.ownsBkImageBitmap = TRUE;
				}
			}
			break;
		}
		case VT_EMPTY:
			if(IsWindow()) {
				LVBKIMAGE bkImage = {0};
				switch(properties.bkImageStyle) {
					case bisWatermark:
						if(RunTimeHelper::IsCommCtrl6()) {
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
							break;
						}
						// not supported - fall through
					case bisNormal:
						bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL;
						break;
					case bisTiled:
						bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_TILE;
						break;
				}
				// OnListSetBkImage will synchronize properties.bkImage
				SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
			} else {
				if(properties.ownsBkImageBitmap && hPreviousBitmap && GetObjectType(hPreviousBitmap) == OBJ_BITMAP) {
					DeleteObject(hPreviousBitmap);
				}
				VariantClear(&properties.bkImage);
			}
			break;
		default:
			return E_FAIL;
			break;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_BKIMAGE);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP ExplorerListView::putref_BkImage(VARIANT newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BKIMAGE);
	switch(newValue.vt) {
		case VT_DISPATCH:
			if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
				LVBKIMAGE bkImage = {0};

				if(newValue.pdispVal) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(newValue.pdispVal);
					if(pPicture) {
						SHORT pictureType = PICTYPE_NONE;
						pPicture->get_Type(&pictureType);
						if(pictureType == PICTYPE_BITMAP) {
							OLE_HANDLE h = NULL;
							pPicture->get_Handle(&h);

							// SysListView32 always needs its own bitmap
							bkImage.hbm = CopyBitmap(static_cast<HBITMAP>(LongToHandle(h)));
						}
					}
				}

				if(bkImage.hbm) {
					bkImage.ulFlags = LVBKIF_SOURCE_HBITMAP;
					if(properties.absoluteBkImagePosition) {
						bkImage.ulFlags |= LVBKIF_FLAG_TILEOFFSET;
					}
					switch(properties.bkImageStyle) {
						case bisNormal:
							bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
							break;
						case bisTiled:
							bkImage.ulFlags |= LVBKIF_STYLE_TILE;
							break;
						case bisWatermark:
							// yes, we've to set LVBKIF_SOURCE_NONE - the MS employee must have been stoned...
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
							break;
					}
					bkImage.xOffsetPercent = properties.bkImagePositionX;
					bkImage.yOffsetPercent = properties.bkImagePositionY;
				} else {
					switch(properties.bkImageStyle) {
						case bisWatermark:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
							break;
						case bisNormal:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL;
							break;
						case bisTiled:
							bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_TILE;
							break;
					}
				}
				// OnListSetBkImage will synchronize properties.bkImage
				SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
				// properties.bkImage now contains a bitmap handle
			} else {
				HBITMAP hPreviousBitmap = NULL;
				VARIANT v;
				VariantInit(&v);
				if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4))) {
					hPreviousBitmap = static_cast<HBITMAP>(LongToHandle(v.lVal));
				}

				HBITMAP hNewBitmap = NULL;
				if(newValue.pdispVal) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(newValue.pdispVal);
					if(pPicture) {
						SHORT pictureType = PICTYPE_NONE;
						pPicture->get_Type(&pictureType);
						if(pictureType == PICTYPE_BITMAP) {
							OLE_HANDLE h = NULL;
							pPicture->get_Handle(&h);

							hNewBitmap = static_cast<HBITMAP>(LongToHandle(h));
						}
					}
				}

				if(properties.ownsBkImageBitmap && hPreviousBitmap && hPreviousBitmap != hNewBitmap && GetObjectType(hPreviousBitmap) == OBJ_BITMAP) {
					DeleteObject(hPreviousBitmap);
				}
				VariantClear(&properties.bkImage);
			}

			properties.bkImage.pdispVal = newValue.pdispVal;
			if(properties.bkImage.pdispVal) {
				properties.bkImage.pdispVal->AddRef();
			}
			properties.bkImage.vt = VT_DISPATCH;
			properties.ownsBkImageBitmap = FALSE;
			break;
		default:
			return E_FAIL;
			break;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_BKIMAGE);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BkImagePositionX(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.bkImagePositionX;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BkImagePositionX(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BKIMAGEPOSITIONX);
	if(IsWindow()) {
		LVBKIMAGE bkImage = {0};
		bkImage.cchImageMax = MAX_PATH;
		TCHAR pBuffer[MAX_PATH + 1];
		bkImage.pszImage = pBuffer;
		SendMessage(LVM_GETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
		if(properties.bkImageStyle == bisWatermark) {
			/* Windows' watermark style implementation is crappy - a listview that has the
			   LVBKIF_TYPE_WATERMARK flag set, won't handle LVM_GETBKIMAGE. */
			bkImage.cchImageMax = 0;
			bkImage.pszImage = NULL;
			bkImage.hbm = flags.hCurrentBackgroundBitmap;
			bkImage.xOffsetPercent = properties.bkImagePositionX;
			bkImage.yOffsetPercent = properties.bkImagePositionY;
			bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
		}

		properties.bkImagePositionX = bkImage.xOffsetPercent = newValue;
		BOOL b = FALSE;
		if(bkImage.hbm) {
			bkImage.cchImageMax = 0;
			bkImage.pszImage = NULL;
			/* SysListView32 destroys its current bitmap if it receives LVM_SETBKIMAGE, so re-assign a copy
			   of the bitmap. */
			HBITMAP hBufferedBitmap = bkImage.hbm;
			bkImage.hbm = CopyBitmap(bkImage.hbm);

			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4)) && hBufferedBitmap == static_cast<HBITMAP>(LongToHandle(v.lVal))) {
				// update property
				VariantClear(&properties.bkImage);
				properties.bkImage.lVal = HandleToLong(bkImage.hbm);
				properties.bkImage.vt = VT_I4;
				b = TRUE;
			}
		}
		// NOTE: We had wParam set to TRUE here. Why?
		SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
		properties.ownsBkImageBitmap = b;
	} else {
		properties.bkImagePositionX = newValue;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_BKIMAGEPOSITIONX);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BkImagePositionY(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.bkImagePositionY;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BkImagePositionY(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BKIMAGEPOSITIONY);
	if(IsWindow()) {
		LVBKIMAGE bkImage = {0};
		bkImage.cchImageMax = MAX_PATH;
		TCHAR pBuffer[MAX_PATH + 1];
		bkImage.pszImage = pBuffer;
		SendMessage(LVM_GETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
		if(properties.bkImageStyle == bisWatermark) {
			/* Windows' watermark style implementation is crappy - a listview that has the
			   LVBKIF_TYPE_WATERMARK flag set, won't handle LVM_GETBKIMAGE. */
			bkImage.cchImageMax = 0;
			bkImage.pszImage = NULL;
			bkImage.hbm = flags.hCurrentBackgroundBitmap;
			bkImage.xOffsetPercent = properties.bkImagePositionX;
			bkImage.yOffsetPercent = properties.bkImagePositionY;
			bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
		}

		BOOL b = FALSE;
		properties.bkImagePositionY = bkImage.yOffsetPercent = newValue;
		if(bkImage.hbm) {
			bkImage.cchImageMax = 0;
			bkImage.pszImage = NULL;
			/* SysListView32 destroys its current bitmap if it receives LVM_SETBKIMAGE, so re-assign a copy
			   of the bitmap. */
			HBITMAP hBufferedBitmap = bkImage.hbm;
			bkImage.hbm = CopyBitmap(bkImage.hbm);

			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4)) && hBufferedBitmap == static_cast<HBITMAP>(LongToHandle(v.lVal))) {
				// update property
				VariantClear(&properties.bkImage);
				properties.bkImage.lVal = HandleToLong(bkImage.hbm);
				properties.bkImage.vt = VT_I4;
				b = TRUE;
			}
		}
		// NOTE: We had wParam set to TRUE here. Why?
		SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
		properties.ownsBkImageBitmap = b;
	} else {
		properties.bkImagePositionY = newValue;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_BKIMAGEPOSITIONY);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BkImageStyle(BkImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BkImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.bkImageStyle;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BkImageStyle(BkImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BKIMAGESTYLE);
	if(properties.bkImageStyle != newValue) {
		BkImageStyleConstants previousStyle = properties.bkImageStyle;
		properties.bkImageStyle = newValue;
		if(IsWindow()) {
			LVBKIMAGE bkImage = {0};
			bkImage.cchImageMax = MAX_PATH;
			TCHAR pBuffer[MAX_PATH + 1];
			bkImage.pszImage = pBuffer;
			SendMessage(LVM_GETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
			if(previousStyle == bisWatermark) {
				/* Windows' watermark style implementation is crappy - a listview that has the
				   LVBKIF_TYPE_WATERMARK flag set, won't handle LVM_GETBKIMAGE. */
				bkImage.cchImageMax = 0;
				bkImage.pszImage = NULL;
				bkImage.hbm = flags.hCurrentBackgroundBitmap;
				bkImage.xOffsetPercent = properties.bkImagePositionX;
				bkImage.yOffsetPercent = properties.bkImagePositionY;
				bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
			}

			BOOL b = FALSE;
			if(bkImage.hbm) {
				bkImage.cchImageMax = 0;
				bkImage.pszImage = NULL;
				/* SysListView32 destroys its current bitmap if it receives LVM_SETBKIMAGE, so re-assign a copy
				   of the bitmap. */
				HBITMAP hBufferedBitmap = bkImage.hbm;
				bkImage.hbm = CopyBitmap(bkImage.hbm);

				VARIANT v;
				VariantInit(&v);
				if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4)) && hBufferedBitmap == static_cast<HBITMAP>(LongToHandle(v.lVal))) {
					// update property
					VariantClear(&properties.bkImage);
					properties.bkImage.lVal = HandleToLong(bkImage.hbm);
					properties.bkImage.vt = VT_I4;
					b = TRUE;
				}
			}

			switch(properties.bkImageStyle) {
				case bisWatermark:
					if(RunTimeHelper::IsCommCtrl6()) {
						/* Windows' watermark style implementation is crappy - the style can't be applied to a
						   listview that already has a background image. */
						LVBKIMAGE dummy = {0};
						dummy.ulFlags = LVBKIF_SOURCE_NONE;
						// NOTE: We had wParam set to TRUE here. Why?
						SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&dummy));

						bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
						break;
					}
					// fall through
				case bisNormal:
					bkImage.ulFlags &= ~(LVBKIF_STYLE_TILE | LVBKIF_TYPE_WATERMARK);
					bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
					break;
				case bisTiled:
					bkImage.ulFlags &= ~(LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK);
					bkImage.ulFlags |= LVBKIF_STYLE_TILE;
					break;
			}

			if(((bkImage.ulFlags & LVBKIF_SOURCE_MASK) == 0) && ((bkImage.ulFlags & LVBKIF_TYPE_WATERMARK) == 0)) {
				// the previous style probably was bisWatermark
				if(bkImage.hbm && RunTimeHelper::IsCommCtrl6()) {
					bkImage.ulFlags |= LVBKIF_SOURCE_HBITMAP;
				} else if(lstrlen(bkImage.pszImage)) {
					bkImage.ulFlags |= LVBKIF_SOURCE_URL;
				}
			}
			// NOTE: We had wParam set to TRUE here. Why?
			SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
			properties.ownsBkImageBitmap = b;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_BKIMAGESTYLE);
	FireViewChange();
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BlendSelectionLasso(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		properties.blendSelectionLasso = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_DOUBLEBUFFER) == LVS_EX_DOUBLEBUFFER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.blendSelectionLasso);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BlendSelectionLasso(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BLENDSELECTIONLASSO);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.blendSelectionLasso != b) {
		properties.blendSelectionLasso = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			if(properties.blendSelectionLasso) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DOUBLEBUFFER, LVS_EX_DOUBLEBUFFER);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DOUBLEBUFFER, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_BLENDSELECTIONLASSO);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BorderSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()/* && IsComctl32Version580OrNewer()*/) {
		properties.borderSelect = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_BORDERSELECT) == LVS_EX_BORDERSELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.borderSelect);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BorderSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BORDERSELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.borderSelect != b) {
		properties.borderSelect = b;
		SetDirty(TRUE);

		if(IsWindow()/* && IsComctl32Version580OrNewer()*/) {
			if(properties.borderSelect) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_BORDERSELECT, LVS_EX_BORDERSELECT);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_BORDERSELECT, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_BORDERSELECT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_CallBackMask(CallBackMaskConstants* pValue)
{
	ATLASSERT_POINTER(pValue, CallBackMaskConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.callBackMask = static_cast<CallBackMaskConstants>(SendMessage(LVM_GETCALLBACKMASK, 0, 0));
	}

	*pValue = properties.callBackMask;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_CallBackMask(CallBackMaskConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_CALLBACKMASK);
	if(properties.callBackMask != newValue) {
		properties.callBackMask = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_SETCALLBACKMASK, properties.callBackMask, 0);
		}
		FireOnChanged(DISPID_EXLVW_CALLBACKMASK);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_CaretColumn(IListViewColumn** ppCaretColumn/* = NULL*/)
{
	ATLASSERT_POINTER(ppCaretColumn, IListViewColumn*);
	if(!ppCaretColumn) {
		return E_POINTER;
	}

	if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		int caretColumn = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETFOCUSEDITEM, 0, 0));
		/*if(caretColumn == 0) {
			// HDM_GETFOCUSEDITEM doesn't distinguish between 'none' and 'column 0'
			caretColumn = -1;
			HDITEM column = {0};
			column.mask = HDI_STATE;
			if(containedSysHeader32.SendMessage(HDM_GETITEM, 0, reinterpret_cast<LPARAM>(&column))) {
				if(column.state & HDIS_FOCUSED) {
					caretColumn = 0;
				}
			}
		}*/
		ClassFactory::InitColumn(caretColumn, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppCaretColumn));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ExplorerListView::putref_CaretColumn(IListViewColumn* pNewCaretColumn)
{
	PUTPROPPROLOG(DISPID_EXLVW_CARETCOLUMN);
	HRESULT hr = E_NOTIMPL;

	if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		hr = E_FAIL;
		int newCaretColumn = -1;
		if(pNewCaretColumn) {
			LONG l = -1;
			pNewCaretColumn->get_Index(&l);
			newCaretColumn = l;
			// TODO: Shouldn't we AddRef' pNewCaretColumn?
		}

		if(containedSysHeader32.SendMessage(HDM_SETFOCUSEDITEM, 0, newCaretColumn)) {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_CARETCOLUMN);
	return hr;
}

STDMETHODIMP ExplorerListView::get_CaretFooterItem(IListViewFooterItem** ppCaretFooterItem)
{
	ATLASSERT_POINTER(ppCaretFooterItem, IListViewFooterItem*);
	if(!ppCaretFooterItem) {
		return E_POINTER;
	}

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		CComPtr<IListViewFooter> pListViewFooter = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListViewFooter), reinterpret_cast<LPARAM>(&pListViewFooter))) {
			ATLASSUME(pListViewFooter);
			int caretFooterItem = -1;
			HRESULT hr = pListViewFooter->GetFooterFocus(&caretFooterItem);
			if(SUCCEEDED(hr)) {
				ClassFactory::InitFooterItem(caretFooterItem, this, IID_IListViewFooterItem, reinterpret_cast<LPUNKNOWN*>(ppCaretFooterItem));
			}
			return hr;
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP ExplorerListView::putref_CaretFooterItem(IListViewFooterItem* pNewCaretFooterItem)
{
	PUTPROPPROLOG(DISPID_EXLVW_CARETFOOTERITEM);
	HRESULT hr = E_NOTIMPL;

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		int newCaretFooterItem = -1;
		if(pNewCaretFooterItem) {
			LONG l = -1;
			pNewCaretFooterItem->get_Index(&l);
			newCaretFooterItem = l;
			// TODO: Shouldn't we AddRef' pNewCaretFooterItem?
		}

		CComPtr<IListViewFooter> pListViewFooter = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListViewFooter), reinterpret_cast<LPARAM>(&pListViewFooter))) {
			ATLASSUME(pListViewFooter);
			// NOTE: Calling SetFooterFocus(-1) doesn't fail, but also doesn't clear the focus.
			return pListViewFooter->SetFooterFocus(newCaretFooterItem);
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_CARETFOOTERITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP ExplorerListView::get_CaretGroup(IListViewGroup** ppCaretGroup)
{
	ATLASSERT_POINTER(ppCaretGroup, IListViewGroup*);
	if(!ppCaretGroup) {
		return E_POINTER;
	}

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		int caretGroup = static_cast<int>(SendMessage(LVM_GETFOCUSEDGROUP, 0, 0));
		ClassFactory::InitGroupByIndex(caretGroup, this, IID_IListViewGroup, reinterpret_cast<LPUNKNOWN*>(ppCaretGroup));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ExplorerListView::putref_CaretGroup(IListViewGroup* pNewCaretGroup)
{
	PUTPROPPROLOG(DISPID_EXLVW_CARETGROUP);
	HRESULT hr = E_NOTIMPL;

	if(IsWindow() && IsComctl32Version610OrNewer()) {
		hr = E_FAIL;
		int newCaretGroup = -1;
		if(pNewCaretGroup) {
			LONG l = -1;
			pNewCaretGroup->get_ID(&l);
			newCaretGroup = l;
			// TODO: Shouldn't we AddRef' pNewCaretGroup?
		}

		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		if(newCaretGroup != -1) {
			group.state = LVGS_FOCUSED;
		}
		group.stateMask = LVGS_FOCUSED;
		group.mask = LVGF_STATE;

		if(SendMessage(LVM_SETGROUPINFO, newCaretGroup, reinterpret_cast<LPARAM>(&group)) == newCaretGroup) {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_CARETGROUP);
	FireViewChange();
	return hr;
}

STDMETHODIMP ExplorerListView::get_CaretItem(VARIANT_BOOL/* changeFocusOnly = VARIANT_FALSE*/, IListViewItem** ppCaretItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppCaretItem, IListViewItem*);
	if(!ppCaretItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		LVITEMINDEX caretItem = {-1, 0};
		if(IsComctl32Version610OrNewer()) {
			SendMessage(LVM_GETNEXTITEMINDEX, reinterpret_cast<WPARAM>(&caretItem), MAKELPARAM(LVNI_FOCUSED, 0));
		} else {
			caretItem.iItem = static_cast<int>(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_FOCUSED, 0)));
		}
		ClassFactory::InitListItem(caretItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppCaretItem));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ExplorerListView::putref_CaretItem(VARIANT_BOOL changeFocusOnly/* = VARIANT_FALSE*/, IListViewItem* pNewCaretItem/* = NULL*/)
{
	PUTPROPPROLOG(DISPID_EXLVW_CARETITEM);
	HRESULT hr = E_FAIL;

	int newCaretItem = -1;
	if(pNewCaretItem) {
		LONG l = -1;
		pNewCaretItem->get_Index(&l);
		newCaretItem = l;
		// TODO: Shouldn't we AddRef' pNewCaretItem?
	}

	if(IsWindow()) {
		LVITEM item = {0};
		if(GetStyle() & LVS_SINGLESEL) {
			if(changeFocusOnly == VARIANT_FALSE) {
				// move selection
				item.stateMask = LVIS_SELECTED;
				int caretItem = static_cast<int>(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_FOCUSED, 0)));
				SendMessage(LVM_SETITEMSTATE, caretItem, reinterpret_cast<LPARAM>(&item));
				if(newCaretItem != -1) {
					// select the 'newCaretItem'
					item.state = LVIS_SELECTED;
				}
			}
		} else {
			if(changeFocusOnly == VARIANT_FALSE) {
				// deselect all items
				/* NOTE: It would be better to deselect all items except the 'newCaretItem', but this would
				         be very slow with many items. */
				item.stateMask = LVIS_SELECTED;
				SendMessage(LVM_SETITEMSTATE, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&item));
				if(newCaretItem != -1) {
					// select the 'newCaretItem'
					item.state = LVIS_SELECTED;
				}
			}
		}

		if(newCaretItem != -1) {
			item.state |= LVIS_FOCUSED;
		}
		item.stateMask |= LVIS_FOCUSED;
		if(SendMessage(LVM_SETITEMSTATE, newCaretItem, reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_CARETITEM);
	FireViewChange();
	return hr;
}

STDMETHODIMP ExplorerListView::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_CheckItemOnSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.checkItemOnSelect = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_AUTOCHECKSELECT) == LVS_EX_AUTOCHECKSELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.checkItemOnSelect);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_CheckItemOnSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_CHECKITEMONSELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.checkItemOnSelect != b) {
		properties.checkItemOnSelect = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.checkItemOnSelect) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOCHECKSELECT, LVS_EX_AUTOCHECKSELECT);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_AUTOCHECKSELECT, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_CHECKITEMONSELECT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ClickableColumnHeaders(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.clickableColumnHeaders = ((containedSysHeader32.GetStyle() & HDS_BUTTONS) == HDS_BUTTONS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.clickableColumnHeaders);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ClickableColumnHeaders(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_CLICKABLECOLUMNHEADERS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.clickableColumnHeaders != b) {
		properties.clickableColumnHeaders = b;
		SetDirty(TRUE);

		if(containedSysHeader32.IsWindow()) {
			// LVS_NOSORTHEADER can't be set after control creation, but the header itself is flexible enough.
			if(properties.clickableColumnHeaders) {
				containedSysHeader32.ModifyStyle(0, HDS_BUTTONS, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				containedSysHeader32.ModifyStyle(HDS_BUTTONS, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			containedSysHeader32.Invalidate();
		}
		if(IsWindow()) {
			// To make the listview style bits match the actual used style, we update the listview, too.
			if(properties.clickableColumnHeaders) {
				ModifyStyle(LVS_NOSORTHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				ModifyStyle(0, LVS_NOSORTHEADER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_CLICKABLECOLUMNHEADERS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ColumnHeaderVisibility(ColumnHeaderVisibilityConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ColumnHeaderVisibilityConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & LVS_NOCOLUMNHEADER) {
			properties.columnHeaderVisibility = chvInvisible;
		} else {
			if(IsComctl32Version610OrNewer() && ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_HEADERINALLVIEWS) == LVS_EX_HEADERINALLVIEWS)) {
				properties.columnHeaderVisibility = chvVisibleInAllViews;
			} else {
				properties.columnHeaderVisibility = chvVisibleInDetailsView;
			}
		}
	}

	*pValue = properties.columnHeaderVisibility;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ColumnHeaderVisibility(ColumnHeaderVisibilityConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_COLUMNHEADERVISIBILITY);
	if(properties.columnHeaderVisibility != newValue) {
		properties.columnHeaderVisibility = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.columnHeaderVisibility) {
				case chvInvisible:
					ModifyStyle(0, LVS_NOCOLUMNHEADER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case chvVisibleInDetailsView:
					ModifyStyle(LVS_NOCOLUMNHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					if(IsComctl32Version610OrNewer()) {
						SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERINALLVIEWS, 0);
					}
					break;
				case chvVisibleInAllViews:
					ModifyStyle(LVS_NOCOLUMNHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					if(IsComctl32Version610OrNewer()) {
						SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERINALLVIEWS, LVS_EX_HEADERINALLVIEWS);
					}
					break;
			}
		}
		FireOnChanged(DISPID_EXLVW_COLUMNHEADERVISIBILITY);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Columns(IListViewColumns** ppColumns)
{
	ATLASSERT_POINTER(ppColumns, IListViewColumns*);
	if(!ppColumns) {
		return E_POINTER;
	}

	ClassFactory::InitColumns(this, IID_IListViewColumns, reinterpret_cast<LPUNKNOWN*>(ppColumns));
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ColumnWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(!IsInViewMode(vList)) {
			return E_FAIL;
		}

		*pValue = static_cast<LONG>(SendMessage(LVM_GETCOLUMNWIDTH, 0, 0));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::put_ColumnWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_COLUMNWIDTH);
	ATLASSERT(newValue > 1);
	if(newValue <= 1) {
		return E_FAIL;
	}

	if(IsInDesignMode()) {
		return E_FAIL;
	}

	if(IsWindow()) {
		if(!IsInViewMode(vDetails)) {
			/* If the control is in 'Icons', 'Small Icons' or 'Tiles' view, it will apply the setting, but
			   will return FALSE. So we don't check the return value here. */
			SendMessage(LVM_SETCOLUMNWIDTH, 0, MAKELPARAM(static_cast<int>(newValue), 0));
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deListMouseEvents) != (newValue & deListMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deListMouseEvents) {
					// nothing to do
				} else {
					if(!cachedSettings.hotTracking) {
						TRACKMOUSEEVENT trackingOptions = {0};
						trackingOptions.cbSize = sizeof(trackingOptions);
						trackingOptions.hwndTrack = *this;
						trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
						TrackMouseEvent(&trackingOptions);
					}
					itemUnderMouse.iItem = -1;
					itemUnderMouse.iGroup = 0;
					subItemUnderMouse = -1;
				}
			}
		}

		if((properties.disabledEvents & deEditMouseEvents) != (newValue & deEditMouseEvents)) {
			if(containedEdit.IsWindow()) {
				if(newValue & deEditMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = containedEdit.m_hWnd;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		if((properties.disabledEvents & deHeaderMouseEvents) != (newValue & deHeaderMouseEvents)) {
			if(containedSysHeader32.IsWindow()) {
				if(newValue & deHeaderMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = containedSysHeader32.m_hWnd;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_EXLVW_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_DragScrollTimeBase(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragScrollTimeBase;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_DragScrollTimeBase(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_DRAGSCROLLTIMEBASE);
	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragScrollTimeBase != newValue) {
		properties.dragScrollTimeBase = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_DRAGSCROLLTIMEBASE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_DraggedColumn(IListViewColumn** ppColumn)
{
	ATLASSERT_POINTER(ppColumn, IListViewColumn*);
	if(!ppColumn) {
		return E_POINTER;
	}

	ClassFactory::InitColumn(dragDropStatus.draggedColumn, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppColumn));
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_DraggedItems(IListViewItemContainer** ppItems)
{
	ATLASSERT_POINTER(ppItems, IListViewItemContainer*);
	if(!ppItems) {
		return E_POINTER;
	}

	*ppItems = NULL;
	if(dragDropStatus.pDraggedItems) {
		return dragDropStatus.pDraggedItems->Clone(ppItems);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_DrawImagesAsynchronously(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.drawImagesAsynchronously = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_DRAWIMAGEASYNC) == LVS_EX_DRAWIMAGEASYNC);
	}

	*pValue = BOOL2VARIANTBOOL(properties.drawImagesAsynchronously);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_DrawImagesAsynchronously(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_DRAWIMAGESASYNCHRONOUSLY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.drawImagesAsynchronously != b) {
		properties.drawImagesAsynchronously = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.drawImagesAsynchronously) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DRAWIMAGEASYNC, LVS_EX_DRAWIMAGEASYNC);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DRAWIMAGEASYNC, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_DRAWIMAGESASYNCHRONOUSLY);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_DropHilitedItem(IListViewItem** ppDropHilitedItem)
{
	ATLASSERT_POINTER(ppDropHilitedItem, IListViewItem*);
	if(!ppDropHilitedItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		LVITEMINDEX dropHilitedItem = {-1, 0};
		if(IsComctl32Version610OrNewer()) {
			SendMessage(LVM_GETNEXTITEMINDEX, reinterpret_cast<WPARAM>(&dropHilitedItem), MAKELPARAM(LVNI_DROPHILITED, 0));
		} else {
			dropHilitedItem.iItem = static_cast<int>(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_DROPHILITED, 0)));
		}
		ClassFactory::InitListItem(dropHilitedItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppDropHilitedItem));
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::putref_DropHilitedItem(IListViewItem* pNewDropHilitedItem)
{
	PUTPROPPROLOG(DISPID_EXLVW_DROPHILITEDITEM);
	HRESULT hr = E_FAIL;

	int newDropHilitedItem = -1;
	if(pNewDropHilitedItem) {
		LONG l = -1;
		pNewDropHilitedItem->get_Index(&l);
		newDropHilitedItem = l;
		// TODO: Shouldn't we AddRef' pNewDropHilitedItem?
	}

	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);

		LVITEM item = {0};
		item.stateMask = LVIS_DROPHILITED;
		// remove the current drop highlight
		if(SendMessage(LVM_SETITEMSTATE, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&item))) {
			hr = S_OK;
		}
		// now set the new one
		if(newDropHilitedItem != -1) {
			item.state = LVIS_DROPHILITED;
			if(SendMessage(LVM_SETITEMSTATE, newDropHilitedItem, reinterpret_cast<LPARAM>(&item))) {
				hr = S_OK;
			}
		}

		dragDropStatus.ShowDragImage(TRUE);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_DROPHILITEDITEM);
	return hr;
}

STDMETHODIMP ExplorerListView::get_EditBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.editBackColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_EditBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EDITBACKCOLOR);
	if(properties.editBackColor != newValue) {
		properties.editBackColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			DeleteObject(hEditBackColorBrush);
			hEditBackColorBrush = CreateSolidBrush(OLECOLOR2COLORREF(properties.editBackColor));
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_EXLVW_EDITBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_EditForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.editForeColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_EditForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EDITFORECOLOR);
	if(properties.editForeColor != newValue) {
		properties.editForeColor = newValue;
		SetDirty(TRUE);

		if(containedEdit.IsWindow()) {
			containedEdit.Invalidate();
		}
		FireOnChanged(DISPID_EXLVW_EDITFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_EditHoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.editHoverTime;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_EditHoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EDITHOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.editHoverTime != newValue) {
		properties.editHoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_EDITHOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_EditIMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.editIMEMode;
	} else {
		if((GetFocus() == containedEdit.m_hWnd) && (GetEffectiveIMEMode_Edit() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(containedEdit.m_hWnd);
			if((ime != imeInherit) && (properties.editIMEMode != imeInherit)) {
				properties.editIMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode_Edit();
	}

	return S_OK;
}

STDMETHODIMP ExplorerListView::put_EditIMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EDITIMEMODE);
	if(properties.editIMEMode != newValue) {
		properties.editIMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == containedEdit.m_hWnd) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode_Edit();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(containedEdit.m_hWnd, ime);
				}
			}
		}
		FireOnChanged(DISPID_EXLVW_EDITIMEMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_EditText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"");
	if(containedEdit.IsWindow()) {
		containedEdit.GetWindowText(pValue);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::put_EditText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EDITTEXT);
	if(!containedEdit.IsWindow()) {
		return E_FAIL;
	}

	containedEdit.SetWindowText(COLE2CT(newValue));

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_EDITTEXT);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_EmptyMarkupText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"");
	/*if(IsComctl32Version610OrNewer()) {
		WCHAR pBuffer[L_MAX_URL_LENGTH + 1];
		ZeroMemory(pBuffer, L_MAX_URL_LENGTH + 1);
		if(SendMessage(LVM_GETEMPTYTEXT, (L_MAX_URL_LENGTH + 1), reinterpret_cast<LPARAM>(pBuffer))) {
			*pValue = _bstr_t(pBuffer).Detach();
		}
	} else {*/
		if(properties.pEmptyMarkupText) {
			*pValue = _bstr_t(properties.pEmptyMarkupText).Detach();
		}
	//}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_EmptyMarkupText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EMPTYMARKUPTEXT);
	if(newValue && lstrlenW(OLE2W(newValue)) >= L_MAX_URL_LENGTH) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(!properties.pEmptyMarkupText || !newValue || (lstrcmpiW(properties.pEmptyMarkupText, OLE2W(newValue)) != 0)) {
		if(properties.pEmptyMarkupText) {
			SECUREFREE(properties.pEmptyMarkupText);
		}
		if(newValue && lstrcmpiW(OLE2W(newValue), L"") != 0) {
			properties.pEmptyMarkupText = reinterpret_cast<LPWSTR>(malloc((L_MAX_URL_LENGTH + 1) * sizeof(WCHAR)));
			if(properties.pEmptyMarkupText) {
				ZeroMemory(properties.pEmptyMarkupText, (L_MAX_URL_LENGTH + 1) * sizeof(WCHAR));
				ATLVERIFY(SUCCEEDED(StringCchCopyW(properties.pEmptyMarkupText, L_MAX_URL_LENGTH, OLE2W(newValue))));
			}
		}
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_RESETEMPTYTEXT, 0, 0);
		}
		FireOnChanged(DISPID_EXLVW_EMPTYMARKUPTEXT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_EmptyMarkupTextAlignment(AlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.emptyMarkupTextAlignment;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_EmptyMarkupTextAlignment(AlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT);
	if(newValue == alRight) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.emptyMarkupTextAlignment != newValue) {
		properties.emptyMarkupTextAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_RESETEMPTYTEXT, 0, 0);
		}
		FireOnChanged(DISPID_EXLVW_EMPTYMARKUPTEXTALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_EXLVW_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_FilterChangedTimeout(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.filterChangedTimeout;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_FilterChangedTimeout(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_FILTERCHANGEDTIMEOUT);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(containedSysHeader32.IsWindow()/* && IsComctl32Version580OrNewer()*/) {
		containedSysHeader32.SendMessage(HDM_SETFILTERCHANGETIMEOUT, 0, (properties.filterChangedTimeout == -1 ? GetDoubleClickTime() * 2 : properties.filterChangedTimeout));
	}
	if(properties.filterChangedTimeout != newValue) {
		properties.filterChangedTimeout = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_FILTERCHANGEDTIMEOUT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_FirstVisibleItem(IListViewItem** ppFirstItem)
{
	ATLASSERT_POINTER(ppFirstItem, IListViewItem*);
	if(!ppFirstItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		LVITEMINDEX firstItem = {-1, 0};
		firstItem.iItem = static_cast<int>(SendMessage(LVM_GETTOPINDEX, 0, 0));
		ClassFactory::InitListItem(firstItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppFirstItem), FALSE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_EXLVW_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_FONT);
	return S_OK;
}

STDMETHODIMP ExplorerListView::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_EXLVW_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_FONT);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_FooterIntroText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"");
	if(properties.footerIntroText.Length() > 0) {
		*pValue = properties.footerIntroText.Copy();
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_FooterIntroText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_FOOTERINTROTEXT);
	if(newValue && lstrlenW(OLE2W(newValue)) >= 4096) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.footerIntroText != newValue) {
		properties.footerIntroText.AssignBSTR(newValue);
		SetDirty(TRUE);
		if(IsWindow() && IsComctl32Version610OrNewer()) {
			CComPtr<IListViewFooter> pListViewFooter = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListViewFooter), reinterpret_cast<LPARAM>(&pListViewFooter))) {
				ATLASSUME(pListViewFooter);
				ATLVERIFY(SUCCEEDED(pListViewFooter->SetIntroText(OLE2W(properties.footerIntroText))));
			}
		}
		FireViewChange();
		FireOnChanged(DISPID_EXLVW_FOOTERINTROTEXT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_FooterItems(IListViewFooterItems** ppFooterItems)
{
	ATLASSERT_POINTER(ppFooterItems, IListViewFooterItems*);
	if(!ppFooterItems) {
		return E_POINTER;
	}

	// we check for comctl32.dll 6.10 here, so we don't have to do in (Virtual)ListViewFooterItem(s)
	if(IsComctl32Version610OrNewer()) {
		ClassFactory::InitFooterItems(this, IID_IListViewFooterItems, reinterpret_cast<LPUNKNOWN*>(ppFooterItems));
	} else {
		*ppFooterItems = NULL;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(LVM_GETTEXTCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.foreColor = 0x80000000 | COLOR_WINDOWTEXT;
		} else if(color != OLECOLOR2COLORREF(properties.foreColor)) {
			properties.foreColor = color;
		}
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_SETTEXTCOLOR, 0, OLECOLOR2COLORREF(properties.foreColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_FullRowSelect(FullRowSelectConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FullRowSelectConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		FullRowSelectConstants fullRowSelect = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_FULLROWSELECT) == LVS_EX_FULLROWSELECT ? frsNormalMode : frsDisabled);
		if(!(fullRowSelect == frsNormalMode && properties.fullRowSelect == frsExtendedMode && IsComctl32Version610OrNewer())) {
			properties.fullRowSelect = fullRowSelect;
		}
	}

	*pValue = properties.fullRowSelect;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_FullRowSelect(FullRowSelectConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_FULLROWSELECT);
	if(newValue == static_cast<FullRowSelectConstants>(VARIANT_TRUE)) {
		newValue = frsNormalMode;
	}
	if(properties.fullRowSelect != newValue) {
		properties.fullRowSelect = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			CComPtr<IListView_WIN7> pListView7 = NULL;
			CComPtr<IListView_WINVISTA> pListViewVista = NULL;
			switch(properties.fullRowSelect) {
				case frsDisabled:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, 0);
					if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
						ATLASSUME(pListView7);
						ATLVERIFY(SUCCEEDED(pListView7->SetSelectionFlags(0x01, 0x00)));
					} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
						ATLASSUME(pListViewVista);
						ATLVERIFY(SUCCEEDED(pListViewVista->SetSelectionFlags(0x01, 0x00)));
					}
					break;
				case frsNormalMode:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, LVS_EX_FULLROWSELECT);
					if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
						ATLASSUME(pListView7);
						ATLVERIFY(SUCCEEDED(pListView7->SetSelectionFlags(0x01, 0x00)));
					} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
						ATLASSUME(pListViewVista);
						ATLVERIFY(SUCCEEDED(pListViewVista->SetSelectionFlags(0x01, 0x00)));
					}
					break;
				case frsExtendedMode:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, LVS_EX_FULLROWSELECT);
					if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
						ATLASSUME(pListView7);
						ATLVERIFY(SUCCEEDED(pListView7->SetSelectionFlags(0x01, 0x01)));
					} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
						ATLASSUME(pListViewVista);
						ATLVERIFY(SUCCEEDED(pListViewVista->SetSelectionFlags(0x01, 0x01)));
					}
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_FULLROWSELECT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GridLines(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.gridLines = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_GRIDLINES) == LVS_EX_GRIDLINES);
	}

	*pValue = BOOL2VARIANTBOOL(properties.gridLines);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GridLines(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GRIDLINES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.gridLines != b) {
		properties.gridLines = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.gridLines) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, LVS_EX_GRIDLINES);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_GRIDLINES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupFooterForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		groupMetrics.mask = LVGMF_TEXTCOLOR;
		SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		if(groupMetrics.crFooter == CLR_NONE) {
			properties.groupFooterForeColor = 0x80000000 | COLOR_WINDOWTEXT;
		} else if(groupMetrics.crFooter != OLECOLOR2COLORREF(properties.groupFooterForeColor)) {
			properties.groupFooterForeColor = groupMetrics.crFooter;
		}
		properties.groupFooterForeColor = groupMetrics.crFooter;
	}

	*pValue = properties.groupFooterForeColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupFooterForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPFOOTERFORECOLOR);
	if(properties.groupFooterForeColor != newValue) {
		properties.groupFooterForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVGROUPMETRICS groupMetrics = {0};
			groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
			groupMetrics.mask = LVGMF_TEXTCOLOR;
			groupMetrics.crFooter = OLECOLOR2COLORREF(properties.groupFooterForeColor);

			SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		}
		FireOnChanged(DISPID_EXLVW_GROUPFOOTERFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupHeaderForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		groupMetrics.mask = LVGMF_TEXTCOLOR;
		SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		if(groupMetrics.crHeader == CLR_NONE) {
			properties.groupHeaderForeColor = 0x80000000 | COLOR_WINDOWTEXT;
		} else if(groupMetrics.crHeader != OLECOLOR2COLORREF(properties.groupHeaderForeColor)) {
			properties.groupHeaderForeColor = groupMetrics.crHeader;
		}
		properties.groupHeaderForeColor = groupMetrics.crHeader;
	}

	*pValue = properties.groupHeaderForeColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupHeaderForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPHEADERFORECOLOR);
	if(properties.groupHeaderForeColor != newValue) {
		properties.groupHeaderForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVGROUPMETRICS groupMetrics = {0};
			groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
			groupMetrics.mask = LVGMF_TEXTCOLOR;
			groupMetrics.crHeader = OLECOLOR2COLORREF(properties.groupHeaderForeColor);

			SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		}
		FireOnChanged(DISPID_EXLVW_GROUPHEADERFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupLocale(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.groupLocale;
	if(*pValue == -1) {
		*pValue = GetThreadLocale();
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupLocale(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPLOCALE);
	properties.groupLocale = newValue;

	FireOnChanged(DISPID_EXLVW_GROUPLOCALE);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupMarginBottom(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		groupMetrics.mask = LVGMF_BORDERSIZE;
		SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		properties.groupMarginBottom = groupMetrics.Bottom;
	}

	*pValue = properties.groupMarginBottom;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupMarginBottom(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPMARGINBOTTOM);
	if(properties.groupMarginBottom != newValue) {
		properties.groupMarginBottom = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			LVGROUPMETRICS groupMetrics = {0};
			groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
			groupMetrics.mask = LVGMF_BORDERSIZE;

			SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
			groupMetrics.Bottom = properties.groupMarginBottom;
			SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		}
		FireOnChanged(DISPID_EXLVW_GROUPMARGINBOTTOM);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupMarginLeft(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		groupMetrics.mask = LVGMF_BORDERSIZE;
		SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		properties.groupMarginLeft = groupMetrics.Left;
	}

	*pValue = properties.groupMarginLeft;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupMarginLeft(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPMARGINLEFT);
	if(properties.groupMarginLeft != newValue) {
		properties.groupMarginLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			LVGROUPMETRICS groupMetrics = {0};
			groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
			groupMetrics.mask = LVGMF_BORDERSIZE;

			SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
			groupMetrics.Left = properties.groupMarginLeft;
			SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		}
		FireOnChanged(DISPID_EXLVW_GROUPMARGINLEFT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupMarginRight(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		groupMetrics.mask = LVGMF_BORDERSIZE;
		SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		properties.groupMarginRight = groupMetrics.Right;
	}

	*pValue = properties.groupMarginRight;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupMarginRight(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPMARGINRIGHT);
	if(properties.groupMarginRight != newValue) {
		properties.groupMarginRight = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			LVGROUPMETRICS groupMetrics = {0};
			groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
			groupMetrics.mask = LVGMF_BORDERSIZE;

			SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
			groupMetrics.Right = properties.groupMarginRight;
			SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		}
		FireOnChanged(DISPID_EXLVW_GROUPMARGINRIGHT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupMarginTop(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		groupMetrics.mask = LVGMF_BORDERSIZE;
		SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		properties.groupMarginTop = groupMetrics.Top;
	}

	*pValue = properties.groupMarginTop;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupMarginTop(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPMARGINTOP);
	if(properties.groupMarginTop != newValue) {
		properties.groupMarginTop = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			LVGROUPMETRICS groupMetrics = {0};
			groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
			groupMetrics.mask = LVGMF_BORDERSIZE;

			SendMessage(LVM_GETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
			groupMetrics.Top = properties.groupMarginTop;
			SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));
		}
		FireOnChanged(DISPID_EXLVW_GROUPMARGINTOP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Groups(IListViewGroups** ppGroups)
{
	ATLASSERT_POINTER(ppGroups, IListViewGroups*);
	if(!ppGroups) {
		return E_POINTER;
	}

	// we check for comctl32.dll 6.0 here, so we don't have to do in (Virtual)ListViewGroup(s)
	if(RunTimeHelper::IsCommCtrl6()) {
		ClassFactory::InitGroups(this, IID_IListViewGroups, reinterpret_cast<LPUNKNOWN*>(ppGroups));
	} else {
		*ppGroups = NULL;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupSortOrder(SortOrderConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SortOrderConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.groupSortOrder;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupSortOrder(SortOrderConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPSORTORDER);
	if(properties.groupSortOrder != newValue) {
		properties.groupSortOrder = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_GROUPSORTORDER);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_GroupTextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(parsingFunction) {
		case tpfCompareString:
			*pValue = properties.groupTextParsingFlagsForCompareString;
			break;
		case tpfVarFromStr:
			*pValue = properties.groupTextParsingFlagsForVarFromStr;
			break;
		default:
			return E_INVALIDARG;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_GroupTextParsingFlags(TextParsingFunctionConstants parsingFunction, LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_GROUPTEXTPARSINGFLAGS);
	switch(parsingFunction) {
		case tpfCompareString:
			properties.groupTextParsingFlagsForCompareString = newValue;
			break;
		case tpfVarFromStr:
			properties.groupTextParsingFlagsForVarFromStr = newValue;
			break;
		default:
			return E_INVALIDARG;
	}

	FireOnChanged(DISPID_EXLVW_GROUPTEXTPARSINGFLAGS);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_hDragImageList(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(dragDropStatus.IsDragging() || dragDropStatus.HeaderIsDragging()) {
		*pValue = HandleToLong(dragDropStatus.hDragImageList);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::get_HeaderFullDragging(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.headerFullDragging = ((containedSysHeader32.GetStyle() & HDS_FULLDRAG) == HDS_FULLDRAG);
	}

	*pValue = BOOL2VARIANTBOOL(properties.headerFullDragging);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HeaderFullDragging(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HEADERFULLDRAGGING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.headerFullDragging != b) {
		properties.headerFullDragging = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.headerFullDragging) {
				containedSysHeader32.ModifyStyle(0, HDS_FULLDRAG, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				containedSysHeader32.ModifyStyle(HDS_FULLDRAG, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
		}
		FireOnChanged(DISPID_EXLVW_HEADERFULLDRAGGING);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HeaderHotTracking(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.headerHotTracking = ((containedSysHeader32.GetStyle() & HDS_HOTTRACK) == HDS_HOTTRACK);
	}

	*pValue = BOOL2VARIANTBOOL(properties.headerHotTracking);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HeaderHotTracking(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HEADERHOTTRACKING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.headerHotTracking != b) {
		properties.headerHotTracking = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.headerHotTracking) {
				containedSysHeader32.ModifyStyle(0, HDS_HOTTRACK, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				containedSysHeader32.ModifyStyle(HDS_HOTTRACK, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
		}
		FireOnChanged(DISPID_EXLVW_HEADERHOTTRACKING);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HeaderHoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.headerHoverTime;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HeaderHoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HEADERHOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.headerHoverTime != newValue) {
		properties.headerHoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_HEADERHOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HeaderOLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.headerOLEDragImageStyle;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HeaderOLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE);
	if(properties.headerOLEDragImageStyle != newValue) {
		properties.headerOLEDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_HEADEROLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HideLabels(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		properties.hideLabels = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_HIDELABELS) == LVS_EX_HIDELABELS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.hideLabels);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HideLabels(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HIDELABELS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hideLabels != b) {
		properties.hideLabels = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			if(properties.hideLabels) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HIDELABELS, LVS_EX_HIDELABELS);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HIDELABELS, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_HIDELABELS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(imageList == ilCurrentView) {
		if(IsInViewMode(vIcons, FALSE)) {
			imageList = ilLarge;
		} else if(IsInViewMode(vTiles, FALSE)) {
			imageList = ilExtraLarge;
		} else {
			imageList = ilSmall;
		}
	}

	*pValue = NULL;
	switch(imageList) {
		case ilSmall:
			if(IsWindow()) {
				*pValue = HandleToLong(cachedSettings.hSmallImageList);
			}
			break;
		case ilLarge:
			if(IsWindow()) {
				*pValue = HandleToLong(cachedSettings.hLargeImageList);
			}
			break;
		case ilExtraLarge:
			if(IsWindow()) {
				*pValue = HandleToLong(cachedSettings.hExtraLargeImageList);
			}
			break;
		case ilHighResolution:
			*pValue = HandleToLong(cachedSettings.hHighResImageList);
			break;
		case ilGroups:
			if(IsWindow() && IsComctl32Version610OrNewer()) {
				*pValue = HandleToLong(cachedSettings.hGroupsImageList);
			}
			break;
		case ilFooterItems:
			if(IsWindow() && IsComctl32Version610OrNewer()) {
				*pValue = HandleToLong(cachedSettings.hGroupsImageList);
			}
			break;
		case ilState:
			if(IsWindow()) {
				*pValue = HandleToLong(cachedSettings.hStateImageList);
			}
			break;
		case ilHeader:
			if(containedSysHeader32.IsWindow()) {
				*pValue = HandleToLong(reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_NORMAL, 0)));
			}
			break;
		case ilHeaderHighResolution:
			*pValue = HandleToLong(cachedSettings.hHighResHeaderImageList);
			break;
		case ilHeaderState:
			if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
				*pValue = HandleToLong(cachedSettings.hHeaderStateImageList);
			}
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HIMAGELIST);
	if(imageList == ilCurrentView) {
		if(IsInViewMode(vIcons, FALSE)) {
			imageList = ilLarge;
		} else if(IsInViewMode(vTiles, FALSE)) {
			imageList = ilExtraLarge;
		} else {
			imageList = ilSmall;
		}
	}

	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilSmall:
			if(IsWindow()) {
				SendMessage(LVM_SETIMAGELIST, LVSIL_SMALL, newValue);
			}
			break;
		case ilLarge:
			if(IsWindow()) {
				if(IsInViewMode(vTiles, FALSE)) {
					cachedSettings.hLargeImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
				} else {
					SendMessage(LVM_SETIMAGELIST, LVSIL_NORMAL, newValue);
				}
			}
			break;
		case ilExtraLarge:
			if(IsWindow()) {
				if(IsInViewMode(vTiles, FALSE)) {
					SendMessage(LVM_SETIMAGELIST, LVSIL_NORMAL, newValue);
				} else {
					cachedSettings.hExtraLargeImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
				}
			}
			break;
		case ilHighResolution:
			cachedSettings.hHighResImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			fireViewChange = FALSE;
			break;
		case ilGroups:
			if(IsWindow() && IsComctl32Version610OrNewer()) {
				SendMessage(LVM_SETIMAGELIST, LVSIL_GROUPHEADER, newValue);
			}
			break;
		case ilFooterItems:
			if(IsWindow() && IsComctl32Version610OrNewer()) {
				SendMessage(LVM_SETIMAGELIST, LVSIL_FOOTERITEMS, newValue);
			}
			break;
		case ilState:
			if(IsWindow()) {
				SendMessage(LVM_SETIMAGELIST, LVSIL_STATE, newValue);
			}
			break;
		case ilHeader:
			if(containedSysHeader32.IsWindow()) {
				containedSysHeader32.SendMessage(HDM_SETIMAGELIST, HDSIL_NORMAL, newValue);
				containedSysHeader32.Invalidate();
			}
			break;
		case ilHeaderHighResolution:
			cachedSettings.hHighResHeaderImageList = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			fireViewChange = FALSE;
			break;
		case ilHeaderState:
			if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
				containedSysHeader32.SendMessage(HDM_SETIMAGELIST, HDSIL_STATE, newValue);
				containedSysHeader32.Invalidate();
			}
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_EXLVW_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HotForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(LVM_GETHOTLIGHTCOLOR, 0, 0));
		if(color == CLR_DEFAULT) {
			properties.hotForeColor = static_cast<OLE_COLOR>(-1);
		} else if(color != OLECOLOR2COLORREF(properties.hotForeColor)) {
			properties.hotForeColor = color;
		}
	}

	*pValue = properties.hotForeColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HotForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTFORECOLOR);

	if(properties.hotForeColor != newValue) {
		properties.hotForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_SETHOTLIGHTCOLOR, 0, (properties.hotForeColor == -1 ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.hotForeColor)));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_HOTFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HotItem(IListViewItem** ppHotItem)
{
	ATLASSERT_POINTER(ppHotItem, IListViewItem*);
	if(!ppHotItem) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		LVITEMINDEX hotItem = {-1, 0};
		CComPtr<IListView_WIN7> pListView7 = NULL;
		CComPtr<IListView_WINVISTA> pListViewVista = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
			ATLASSUME(pListView7);
			hr = pListView7->GetHotItem(&hotItem);
		} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
			ATLASSUME(pListViewVista);
			hr = pListViewVista->GetHotItem(&hotItem);
		} else {
			hotItem.iItem = static_cast<int>(SendMessage(LVM_GETHOTITEM, 0, 0));
			hr = S_OK;
		}
		ClassFactory::InitListItem(hotItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppHotItem));
	}
	return hr;
}

STDMETHODIMP ExplorerListView::putref_HotItem(IListViewItem* pNewHotItem)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTITEM);
	HRESULT hr = E_FAIL;

	LVITEMINDEX newHotItem = {-1, 0};
	if(pNewHotItem) {
		// TODO: fully support LVITEMINDEX
		LONG l = -1;
		pNewHotItem->get_Index(&l);
		newHotItem.iItem = l;
		// TODO: Shouldn't we AddRef' pNewHotItem?
	}

	if(IsWindow()) {
		CComPtr<IListView_WIN7> pListView7 = NULL;
		CComPtr<IListView_WINVISTA> pListViewVista = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
			ATLASSUME(pListView7);
			hr = pListView7->SetHotItem(newHotItem, NULL);
		} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
			ATLASSUME(pListViewVista);
			hr = pListViewVista->SetHotItem(newHotItem, NULL);
		} else if(SendMessage(LVM_SETHOTITEM, newHotItem.iItem, 0)) {
			hr = S_OK;
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_HOTITEM);
	return hr;
}

STDMETHODIMP ExplorerListView::get_HotMouseIcon(IPictureDisp** ppHotMouseIcon)
{
	ATLASSERT_POINTER(ppHotMouseIcon, IPictureDisp*);
	if(!ppHotMouseIcon) {
		return E_POINTER;
	}

	*ppHotMouseIcon = properties.hotMouseIcon.pPictureDisp;
	if(*ppHotMouseIcon) {
		(*ppHotMouseIcon)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HotMouseIcon(IPictureDisp* pNewHotMouseIcon)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTMOUSEICON);
	if(properties.hotMouseIcon.pPictureDisp != pNewHotMouseIcon) {
		properties.hotMouseIcon.StopWatching();
		properties.hotMouseIcon.pPictureDisp = NULL;
		if(pNewHotMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewHotMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.hotMouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.hotMouseIcon.StartWatching();

		if(IsWindow() && (properties.hotMousePointer == mpCustom)) {
			if(properties.hotMouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.hotMouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE hCursor = NULL;
					pPicture->get_Handle(&hCursor);
					properties.hotMouseIcon.dontGetPictureObject = TRUE;
					SendMessage(LVM_SETHOTCURSOR, 0, reinterpret_cast<LPARAM>(LongToHandle(hCursor)));
					properties.hotMouseIcon.dontGetPictureObject = FALSE;
				}
			}
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_HOTMOUSEICON);
	return S_OK;
}

STDMETHODIMP ExplorerListView::putref_HotMouseIcon(IPictureDisp* pNewHotMouseIcon)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTMOUSEICON);
	if(properties.hotMouseIcon.pPictureDisp != pNewHotMouseIcon) {
		properties.hotMouseIcon.StopWatching();
		properties.hotMouseIcon.pPictureDisp = pNewHotMouseIcon;
		properties.hotMouseIcon.StartWatching();

		if(IsWindow() && (properties.hotMousePointer == mpCustom)) {
			if(properties.hotMouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.hotMouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE hCursor = NULL;
					pPicture->get_Handle(&hCursor);
					properties.hotMouseIcon.dontGetPictureObject = TRUE;
					SendMessage(LVM_SETHOTCURSOR, 0, reinterpret_cast<LPARAM>(LongToHandle(hCursor)));
					properties.hotMouseIcon.dontGetPictureObject = FALSE;
				}
			}
		}
	} else if(pNewHotMouseIcon) {
		pNewHotMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_HOTMOUSEICON);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HotMousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hotMousePointer;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HotMousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTMOUSEPOINTER);
	if(properties.hotMousePointer != newValue) {
		properties.hotMousePointer = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.hotMousePointer == mpCustom) {
				if(properties.hotMouseIcon.pPictureDisp) {
					CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.hotMouseIcon.pPictureDisp);
					if(pPicture) {
						OLE_HANDLE hCursor = NULL;
						pPicture->get_Handle(&hCursor);
						properties.hotMouseIcon.dontGetPictureObject = TRUE;
						SendMessage(LVM_SETHOTCURSOR, 0, reinterpret_cast<LPARAM>(LongToHandle(hCursor)));
						properties.hotMouseIcon.dontGetPictureObject = FALSE;
					}
				}
			} else if(properties.hotMousePointer != mpDefault) {
				HCURSOR hCursor = MousePointerConst2hCursor(properties.hotMousePointer);
				properties.hotMouseIcon.dontGetPictureObject = TRUE;
				SendMessage(LVM_SETHOTCURSOR, 0, reinterpret_cast<LPARAM>(hCursor));
				properties.hotMouseIcon.dontGetPictureObject = FALSE;
			} else {
				// reset to default
				properties.hotMouseIcon.dontGetPictureObject = TRUE;
				SendMessage(LVM_SETHOTCURSOR, 0, 0);
				properties.hotMouseIcon.dontGetPictureObject = FALSE;
			}
		}
		FireOnChanged(DISPID_EXLVW_HOTMOUSEPOINTER);
	}

	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HotTracking(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.hotTracking = cachedSettings.hotTracking;
	}

	*pValue = BOOL2VARIANTBOOL(properties.hotTracking);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HotTracking(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTTRACKING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hotTracking != b) {
		properties.hotTracking = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.hotTracking) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT, LVS_EX_TRACKSELECT);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_HOTTRACKING);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HotTrackingHoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.hotTrackingHoverTime = static_cast<LONG>(SendMessage(LVM_GETHOVERTIME, 0, 0));
	}

	*pValue = properties.hotTrackingHoverTime;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HotTrackingHoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOTTRACKINGHOVERTIME);
	if((newValue <= 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hotTrackingHoverTime != newValue) {
		properties.hotTrackingHoverTime = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_SETHOVERTIME, 0, properties.hotTrackingHoverTime);
		}
		FireOnChanged(DISPID_EXLVW_HOTTRACKINGHOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_hWndEdit(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		//*pValue = HandleToLong(static_cast<HWND>(SendMessage(LVM_GETEDITCONTROL, 0, 0)));
		*pValue = HandleToLong(containedEdit.m_hWnd);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_hWndHeader(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		//*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(LVM_GETHEADER, 0, 0)));
		*pValue = HandleToLong(containedSysHeader32.m_hWnd);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(LVM_GETTOOLTIPS, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_HWNDTOOLTIP);
	if(IsWindow()) {
		SendMessage(LVM_SETTOOLTIPS, 0, newValue);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_IMEMode(IMEModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, IMEModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsInDesignMode()) {
		*pValue = properties.IMEMode;
	} else {
		if((GetFocus() == *this) && (GetEffectiveIMEMode() != imeNoControl)) {
			// we have control over the IME, so retrieve its current config
			IMEModeConstants ime = GetCurrentIMEContextMode(*this);
			if((ime != imeInherit) && (properties.IMEMode != imeInherit)) {
				properties.IMEMode = ime;
			}
		}
		*pValue = GetEffectiveIMEMode();
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_IMEMode(IMEModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_IMEMODE);
	if(properties.IMEMode != newValue) {
		properties.IMEMode = newValue;
		SetDirty(TRUE);

		if(!IsInDesignMode()) {
			if(GetFocus() == *this) {
				// we have control over the IME, so update its config
				IMEModeConstants ime = GetEffectiveIMEMode();
				if(ime != imeNoControl) {
					SetCurrentIMEContextMode(*this, ime);
				}
			} else if((GetFocus() == containedEdit.m_hWnd) && (properties.editIMEMode == imeInherit)) {
				// the contained edit control inherits from us and has the focus, so update the IME's config
				SetCurrentIMEContextMode(containedEdit.m_hWnd, GetEffectiveIMEMode_Edit());
			}
		}
		FireOnChanged(DISPID_EXLVW_IMEMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_IncludeHeaderInTabOrder(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.includeHeaderInTabOrder);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_IncludeHeaderInTabOrder(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_INCLUDEHEADERINTABORDER);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.includeHeaderInTabOrder != b) {
		properties.includeHeaderInTabOrder = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_INCLUDEHEADERINTABORDER);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_IncrementalSearchString(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsWindow()) {
		return E_FAIL;
	}

	*pValue = SysAllocString(L"");
	UINT chars = static_cast<UINT>(SendMessage(LVM_GETISEARCHSTRING, 0, 0));
	if(chars > 0) {
		TCHAR* pBuffer = new TCHAR[chars + 1];
		SendMessage(LVM_GETISEARCHSTRING, 0, reinterpret_cast<LPARAM>(pBuffer));
		*pValue = _bstr_t(pBuffer).Detach();
		delete[] pBuffer;
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_InsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(LVM_GETINSERTMARKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.insertMarkColor = 0;
		} else if(color != OLECOLOR2COLORREF(properties.insertMarkColor)) {
			properties.insertMarkColor = color;
		}
	}

	*pValue = properties.insertMarkColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_InsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_INSERTMARKCOLOR);
	if(properties.insertMarkColor != newValue) {
		properties.insertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(LVM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_INSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ItemActivationMode(ItemActivationModeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemActivationModeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0));
		if(style & LVS_EX_ONECLICKACTIVATE) {
			properties.itemActivationMode = iamOneSingleClick;
		} else if(style & LVS_EX_TWOCLICKACTIVATE) {
			properties.itemActivationMode = iamTwoSingleClicks;
		} else {
			properties.itemActivationMode = iamOneDoubleClick;
		}
	}

	*pValue = properties.itemActivationMode;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ItemActivationMode(ItemActivationModeConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ITEMACTIVATIONMODE);
	if(properties.itemActivationMode != newValue) {
		properties.itemActivationMode = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.itemActivationMode) {
				case iamOneSingleClick:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, (LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE), LVS_EX_ONECLICKACTIVATE);
					break;
				case iamTwoSingleClicks:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, (LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE), LVS_EX_TWOCLICKACTIVATE);
					break;
				case iamOneDoubleClick:
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, (LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE), 0);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_ITEMACTIVATIONMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ItemAlignment(ItemAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		/* FIXME: We've problems in 'Icons', 'Small Icons' and 'Tiles' view, because
		          LVS_OWNERDRAWFIXED = LVS_ALIGNBOTTOM */
		#ifdef ALLOWBOTTOMRIGHTALIGNMENT
			switch(GetStyle() & (LVS_ALIGNTOP | LVS_ALIGNLEFT | LVS_ALIGNBOTTOM | LVS_ALIGNRIGHT)) {
				case LVS_ALIGNTOP:
					properties.itemAlignment = iaTop;
					break;
				case LVS_ALIGNLEFT:
					properties.itemAlignment = iaLeft;
					break;
				case LVS_ALIGNBOTTOM:
					properties.itemAlignment = iaBottom;
					break;
				case LVS_ALIGNRIGHT:
					properties.itemAlignment = iaRight;
					break;
			}
		#else
			properties.itemAlignment = ((GetStyle() & LVS_ALIGNLEFT) == LVS_ALIGNLEFT ? iaLeft : iaTop);
		#endif
	}

	*pValue = properties.itemAlignment;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ItemAlignment(ItemAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ITEMALIGNMENT);
	if(properties.itemAlignment != newValue) {
		properties.itemAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			/* FIXME: We've problems in 'Icons', 'Small Icons' and 'Tiles' view, because
			          LVS_OWNERDRAWFIXED = LVS_ALIGNBOTTOM */
			#ifdef ALLOWBOTTOMRIGHTALIGNMENT
				switch(properties.itemAlignment) {
					case iaTop:
						ModifyStyle(LVS_ALIGNLEFT | LVS_ALIGNBOTTOM | LVS_ALIGNRIGHT, LVS_ALIGNTOP);
						break;
					case iaLeft:
						ModifyStyle(LVS_ALIGNTOP | LVS_ALIGNBOTTOM | LVS_ALIGNRIGHT, LVS_ALIGNLEFT);
						break;
					case iaBottom:
						ModifyStyle(LVS_ALIGNTOP | LVS_ALIGNLEFT | LVS_ALIGNRIGHT, LVS_ALIGNBOTTOM);
						break;
					case iaRight:
						ModifyStyle(LVS_ALIGNTOP | LVS_ALIGNLEFT | LVS_ALIGNBOTTOM, LVS_ALIGNRIGHT);
						break;
				}
			#else
				switch(properties.itemAlignment) {
					case iaTop:
						ModifyStyle(LVS_ALIGNLEFT, LVS_ALIGNTOP);
						break;
					case iaLeft:
						ModifyStyle(LVS_ALIGNTOP, LVS_ALIGNLEFT);
						break;
				}
			#endif

			// temporarily change view mode to make the listview update itself
			if(RunTimeHelper::IsCommCtrl6()) {
				switch(SendMessage(LVM_GETVIEW, 0, 0)) {
					case LV_VIEW_ICON:
						SendMessage(LVM_SETVIEW, LV_VIEW_SMALLICON, 0);
						SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
						break;
					case LV_VIEW_SMALLICON:
						SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
						SendMessage(LVM_SETVIEW, LV_VIEW_SMALLICON, 0);
						break;
					case LV_VIEW_TILE:
						SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
						SendMessage(LVM_SETVIEW, LV_VIEW_TILE, 0);
						break;
				}
			} else {
				switch(GetStyle() & LVS_TYPEMASK) {
					case LVS_ICON:
						ModifyStyle(LVS_ICON | LVS_LIST | LVS_REPORT, LVS_SMALLICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						ModifyStyle(LVS_SMALLICON | LVS_LIST | LVS_REPORT, LVS_ICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
					case LVS_SMALLICON:
						ModifyStyle(LVS_SMALLICON | LVS_LIST | LVS_REPORT, LVS_ICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						ModifyStyle(LVS_ICON | LVS_LIST | LVS_REPORT, LVS_SMALLICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
				}
			}
		}
		FireOnChanged(DISPID_EXLVW_ITEMALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ItemBoundingBoxDefinition(ItemBoundingBoxDefinitionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ItemBoundingBoxDefinitionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemBoundingBoxDefinition;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ItemBoundingBoxDefinition(ItemBoundingBoxDefinitionConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ITEMBOUNDINGBOXDEFINITION);
	if(properties.itemBoundingBoxDefinition != newValue) {
		properties.itemBoundingBoxDefinition = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_ITEMBOUNDINGBOXDEFINITION);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ItemHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.itemHeight;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ItemHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_ITEMHEIGHT);
	if(properties.itemHeight != newValue) {
		properties.itemHeight = newValue;
		SetDirty(TRUE);
		if(IsWindow()) {
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_ITEMHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_JustifyIconColumns(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.justifyIconColumns = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_JUSTIFYCOLUMNS) == LVS_EX_JUSTIFYCOLUMNS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.justifyIconColumns);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_JustifyIconColumns(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_JUSTIFYICONCOLUMNS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.justifyIconColumns != b) {
		properties.justifyIconColumns = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.justifyIconColumns) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_JUSTIFYCOLUMNS, LVS_EX_JUSTIFYCOLUMNS);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_JUSTIFYCOLUMNS, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_JUSTIFYICONCOLUMNS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_LabelWrap(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.labelWrap = ((GetStyle() & LVS_NOLABELWRAP) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.labelWrap);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_LabelWrap(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_LABELWRAP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.labelWrap != b) {
		properties.labelWrap = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.labelWrap) {
				ModifyStyle(LVS_NOLABELWRAP, 0);
			} else {
				ModifyStyle(0, LVS_NOLABELWRAP);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_LABELWRAP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ListItems(IListViewItems** ppItems)
{
	ATLASSERT_POINTER(ppItems, IListViewItems*);
	if(!ppItems) {
		return E_POINTER;
	}

	ClassFactory::InitListItems(this, IID_IListViewItems, reinterpret_cast<LPUNKNOWN*>(ppItems));
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_MinItemRowsVisibleInGroups(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		CComPtr<IListView_WIN7> pListView7 = NULL;
		CComPtr<IListView_WINVISTA> pListViewVista = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
			ATLASSUME(pListView7);
			int count = 0;
			ATLVERIFY(SUCCEEDED(pListView7->GetGroupSubsetCount(&count)));
			properties.minItemRowsVisibleInGroups = count;
		} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
			ATLASSUME(pListViewVista);
			int count = 0;
			ATLVERIFY(SUCCEEDED(pListViewVista->GetGroupSubsetCount(&count)));
			properties.minItemRowsVisibleInGroups = count;
		}
	}

	*pValue = properties.minItemRowsVisibleInGroups;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_MinItemRowsVisibleInGroups(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_MINITEMROWSVISIBLEINGROUPS);

	if(properties.minItemRowsVisibleInGroups != newValue) {
		properties.minItemRowsVisibleInGroups = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			CComPtr<IListView_WIN7> pListView7 = NULL;
			CComPtr<IListView_WINVISTA> pListViewVista = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
				ATLASSUME(pListView7);
				ATLVERIFY(SUCCEEDED(pListView7->SetGroupSubsetCount(properties.minItemRowsVisibleInGroups)));
				FireViewChange();
			} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
				ATLASSUME(pListViewVista);
				ATLVERIFY(SUCCEEDED(pListViewVista->SetGroupSubsetCount(properties.minItemRowsVisibleInGroups)));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXLVW_MINITEMROWSVISIBLEINGROUPS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_EXLVW_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ExplorerListView::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_EXLVW_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_MultiSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiSelect = ((GetStyle() & LVS_SINGLESEL) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.multiSelect);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_MultiSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_MULTISELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiSelect != b) {
		properties.multiSelect = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.multiSelect) {
				ModifyStyle(LVS_SINGLESEL, 0);
			} else {
				ModifyStyle(0, LVS_SINGLESEL);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_MULTISELECT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_OutlineColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(LVM_GETOUTLINECOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.outlineColor = 0x80000000 | COLOR_BTNFACE;
		} else if(color != OLECOLOR2COLORREF(properties.outlineColor)) {
			properties.outlineColor = color;
		}
	}

	*pValue = properties.outlineColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_OutlineColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_OUTLINECOLOR);
	if(properties.outlineColor != newValue) {
		properties.outlineColor = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(LVM_SETOUTLINECOLOR, 0, OLECOLOR2COLORREF(properties.outlineColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_OUTLINECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_OwnerDrawn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		/* FIXME: We've problems in 'Icons', 'Small Icons' and 'Tiles' view, because
		          LVS_OWNERDRAWFIXED = LVS_ALIGNBOTTOM */
		properties.ownerDrawn = ((GetStyle() & LVS_OWNERDRAWFIXED) == LVS_OWNERDRAWFIXED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.ownerDrawn);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_OwnerDrawn(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_OWNERDRAWN);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.ownerDrawn != b) {
		properties.ownerDrawn = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			/* FIXME: We've problems in 'Icons', 'Small Icons' and 'Tiles' view, because
			          LVS_OWNERDRAWFIXED = LVS_ALIGNBOTTOM */
			if(properties.ownerDrawn) {
				ModifyStyle(0, LVS_OWNERDRAWFIXED);
			} else {
				ModifyStyle(LVS_OWNERDRAWFIXED, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_OWNERDRAWN);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Regional(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.regional = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_REGIONAL) == LVS_EX_REGIONAL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.regional);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_Regional(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_REGIONAL);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.regional != b) {
		properties.regional = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.regional) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_REGIONAL, LVS_EX_REGIONAL);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_REGIONAL, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_REGIONAL);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_EXLVW_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ResizableColumns(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		properties.resizableColumns = ((containedSysHeader32.GetStyle() & HDS_NOSIZING) == 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.resizableColumns);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ResizableColumns(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_RESIZABLECOLUMNS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.resizableColumns != b) {
		properties.resizableColumns = b;
		SetDirty(TRUE);

		if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.resizableColumns) {
				containedSysHeader32.ModifyStyle(HDS_NOSIZING, 0);
			} else {
				containedSysHeader32.ModifyStyle(0, HDS_NOSIZING);
			}
		}
		FireOnChanged(DISPID_EXLVW_RESIZABLECOLUMNS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetExStyle();
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			BOOL requiresRecreation = FALSE;
			if(properties.rightToLeft & rtlText) {
				if(!(GetExStyle() & WS_EX_RTLREADING)) {
					requiresRecreation = TRUE;
				}
			} else {
				if(GetExStyle() & WS_EX_RTLREADING) {
					requiresRecreation = TRUE;
				}
			}
			if(requiresRecreation) {
				RecreateControlWindow();
			} else {
				if(properties.rightToLeft & rtlLayout) {
					ModifyStyleEx(0, WS_EX_LAYOUTRTL);
				} else {
					ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
				}
				FireViewChange();
			}

			if(containedSysHeader32.IsWindow()) {
				if(GetExStyle() & WS_EX_RTLREADING) {
					containedSysHeader32.ModifyStyleEx(0, WS_EX_RTLREADING, SWP_DRAWFRAME | SWP_FRAMECHANGED);

					HDITEM column = {0};
					column.mask = HDI_FORMAT;
					int columns = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0));
					for(int columnIndex = 0; columnIndex < columns; ++columnIndex) {
						containedSysHeader32.SendMessage(HDM_GETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
						column.fmt |= HDF_RTLREADING;
						containedSysHeader32.SendMessage(HDM_SETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
					}
				} else {
					containedSysHeader32.ModifyStyleEx(WS_EX_RTLREADING, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);

					HDITEM column = {0};
					column.mask = HDI_FORMAT;
					int columns = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0));
					for(int columnIndex = 0; columnIndex < columns; ++columnIndex) {
						containedSysHeader32.SendMessage(HDM_GETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
						column.fmt &= ~HDF_RTLREADING;
						containedSysHeader32.SendMessage(HDM_SETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
					}
				}
				if(GetExStyle() & WS_EX_LAYOUTRTL) {
					containedSysHeader32.ModifyStyleEx(0, WS_EX_LAYOUTRTL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				} else {
					containedSysHeader32.ModifyStyleEx(WS_EX_LAYOUTRTL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				}
				containedSysHeader32.Invalidate();
			}
		}
		FireOnChanged(DISPID_EXLVW_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ScrollBars(ScrollBarsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollBarsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & LVS_NOSCROLL) {
			properties.scrollBars = sbNone;
		} else if((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_FLATSB) == LVS_EX_FLATSB) {
			properties.scrollBars = sbFlat;
		} else {
			properties.scrollBars = sbNormal;
		}
	}

	*pValue = properties.scrollBars;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ScrollBars(ScrollBarsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SCROLLBARS);
	if(properties.scrollBars != newValue) {
		properties.scrollBars = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			BOOL requiresRecreation = FALSE;
			switch(properties.scrollBars) {
				case sbNone:
					requiresRecreation = TRUE;
					/*ModifyStyle(0, LVS_NOSCROLL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FLATSB, 0);*/
					break;
				case sbNormal:
					if((GetStyle() & LVS_NOSCROLL) == LVS_NOSCROLL) {
						requiresRecreation = TRUE;
					} else {
						ModifyStyle(LVS_NOSCROLL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FLATSB, 0);
					}
					break;
				case sbFlat:
					if((GetStyle() & LVS_NOSCROLL) == LVS_NOSCROLL) {
						requiresRecreation = TRUE;
					} else {
						ModifyStyle(LVS_NOSCROLL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FLATSB, LVS_EX_FLATSB);
					}
					break;
			}
			if(requiresRecreation) {
				RecreateControlWindow();
			} else {
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXLVW_SCROLLBARS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_SelectedColumn(IListViewColumn** ppSelectedColumn)
{
	ATLASSERT_POINTER(ppSelectedColumn, IListViewColumn*);
	if(!ppSelectedColumn) {
		return E_POINTER;
	}

	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		int selectedColumn = static_cast<int>(SendMessage(LVM_GETSELECTEDCOLUMN, 0, 0));
		ClassFactory::InitColumn(selectedColumn, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppSelectedColumn));
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ExplorerListView::putref_SelectedColumn(IListViewColumn* pNewSelectedColumn)
{
	PUTPROPPROLOG(DISPID_EXLVW_SELECTEDCOLUMN);
	HRESULT hr = E_FAIL;

	int newSelectedColumn = -1;
	if(pNewSelectedColumn) {
		LONG l = -1;
		pNewSelectedColumn->get_Index(&l);
		newSelectedColumn = l;
		// TODO: Shouldn't we AddRef' pNewSelectedColumn?
	}

	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		SendMessage(LVM_SETSELECTEDCOLUMN, newSelectedColumn, 0);
		FireViewChange();
		hr = S_OK;
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_EXLVW_SELECTEDCOLUMN);
	return hr;
}

STDMETHODIMP ExplorerListView::get_SelectedColumnBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		CComPtr<IVisualProperties> pVisualProperties = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IVisualProperties), reinterpret_cast<LPARAM>(&pVisualProperties))) {
			ATLASSUME(pVisualProperties);
			COLORREF color;
			ATLVERIFY(SUCCEEDED(pVisualProperties->GetColor(VPCF_SORTCOLUMN, &color)));
			if(color == properties.defaultSelectedColumnBackColor) {
				properties.selectedColumnBackColor = static_cast<OLE_COLOR>(-1);
			} else if(color != OLECOLOR2COLORREF(properties.selectedColumnBackColor)) {
				properties.selectedColumnBackColor = color;
			}
		}
	}

	*pValue = properties.selectedColumnBackColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_SelectedColumnBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SELECTEDCOLUMNBACKCOLOR);
	if(properties.selectedColumnBackColor != newValue) {
		properties.selectedColumnBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			CComPtr<IVisualProperties> pVisualProperties = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IVisualProperties), reinterpret_cast<LPARAM>(&pVisualProperties))) {
				ATLASSUME(pVisualProperties);
				ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SORTCOLUMN, (properties.selectedColumnBackColor == -1 ? properties.defaultSelectedColumnBackColor : OLECOLOR2COLORREF(properties.selectedColumnBackColor)))));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXLVW_SELECTEDCOLUMNBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!(dragDropStatus.IsDragging() || dragDropStatus.HeaderIsDragging())) {
		return E_FAIL;
	}

	*pValue = BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWDRAGIMAGE);
	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_EXLVW_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowFilterBar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedSysHeader32.IsWindow()/* && IsComctl32Version580OrNewer()*/) {
		properties.showFilterBar = ((containedSysHeader32.GetStyle() & HDS_FILTERBAR) == HDS_FILTERBAR);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showFilterBar);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowFilterBar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWFILTERBAR);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showFilterBar != b) {
		properties.showFilterBar = b;
		SetDirty(TRUE);

		if(containedSysHeader32.IsWindow()/* && IsComctl32Version580OrNewer()*/) {
			if(properties.showFilterBar) {
				containedSysHeader32.ModifyStyle(0, HDS_FILTERBAR, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				containedSysHeader32.ModifyStyle(HDS_FILTERBAR, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			containedSysHeader32.Invalidate();

			if(IsWindow()) {
				// the header won't resize automatically, so force a resize by temporarily hiding it
				if(GetStyle() & LVS_NOCOLUMNHEADER) {
					ModifyStyle(LVS_NOCOLUMNHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					ModifyStyle(0, LVS_NOCOLUMNHEADER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				} else {
					ModifyStyle(0, LVS_NOCOLUMNHEADER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					ModifyStyle(LVS_NOCOLUMNHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				}
			}
		}
		FireOnChanged(DISPID_EXLVW_SHOWFILTERBAR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowGroups(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		properties.showGroups = (SendMessage(LVM_ISGROUPVIEWENABLED, 0, 0) == TRUE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showGroups);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowGroups(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWGROUPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showGroups != b) {
		properties.showGroups = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(LVM_ENABLEGROUPVIEW, properties.showGroups, 0);
		}
		FireOnChanged(DISPID_EXLVW_SHOWGROUPS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowHeaderChevron(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		//properties.showHeaderChevron = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_COLUMNOVERFLOW) == LVS_EX_COLUMNOVERFLOW);
		properties.showHeaderChevron = ((containedSysHeader32.GetStyle() & HDS_OVERFLOW) == HDS_OVERFLOW);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showHeaderChevron);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowHeaderChevron(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWHEADERCHEVRON);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showHeaderChevron != b) {
		properties.showHeaderChevron = b;
		SetDirty(TRUE);

		if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
			/*if(properties.showHeaderChevron) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_COLUMNOVERFLOW, LVS_EX_COLUMNOVERFLOW);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_COLUMNOVERFLOW, 0);
			}
			FireViewChange();*/
			if(properties.showHeaderChevron) {
				containedSysHeader32.ModifyStyle(0, HDS_OVERFLOW);
			} else {
				containedSysHeader32.ModifyStyle(HDS_OVERFLOW, 0);
			}
			containedSysHeader32.Invalidate();
		}
		FireOnChanged(DISPID_EXLVW_SHOWHEADERCHEVRON);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowHeaderStateImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		properties.showHeaderStateImages = ((containedSysHeader32.GetStyle() & HDS_CHECKBOXES) == HDS_CHECKBOXES);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showHeaderStateImages);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowHeaderStateImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWHEADERSTATEIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showHeaderStateImages != b) {
		properties.showHeaderStateImages = b;
		SetDirty(TRUE);

		if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.showHeaderStateImages) {
				containedSysHeader32.ModifyStyle(0, HDS_CHECKBOXES);
			} else {
				containedSysHeader32.ModifyStyle(HDS_CHECKBOXES, 0);
			}
			containedSysHeader32.Invalidate();
		}
		FireOnChanged(DISPID_EXLVW_SHOWHEADERSTATEIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowStateImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showStateImages = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_CHECKBOXES) == LVS_EX_CHECKBOXES);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showStateImages);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowStateImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWSTATEIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showStateImages != b) {
		properties.showStateImages = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showStateImages) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, LVS_EX_CHECKBOXES);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_SHOWSTATEIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ShowSubItemImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showSubItemImages = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_SUBITEMIMAGES) == LVS_EX_SUBITEMIMAGES);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showSubItemImages);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ShowSubItemImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SHOWSUBITEMIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showSubItemImages != b) {
		properties.showSubItemImages = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showSubItemImages) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SUBITEMIMAGES, LVS_EX_SUBITEMIMAGES);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SUBITEMIMAGES, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_SHOWSUBITEMIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_SimpleSelect(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		properties.simpleSelect = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_SIMPLESELECT) == LVS_EX_SIMPLESELECT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.simpleSelect);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_SimpleSelect(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SIMPLESELECT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.simpleSelect != b) {
		properties.simpleSelect = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			if(properties.simpleSelect) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SIMPLESELECT, LVS_EX_SIMPLESELECT);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SIMPLESELECT, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_SIMPLESELECT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_SingleRow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		properties.singleRow = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_SINGLEROW) == LVS_EX_SINGLEROW);
	}

	*pValue = BOOL2VARIANTBOOL(properties.singleRow);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_SingleRow(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SINGLEROW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.singleRow != b) {
		properties.singleRow = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			if(properties.singleRow) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SINGLEROW, LVS_EX_SINGLEROW);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SINGLEROW, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_SINGLEROW);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_SnapToGrid(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		properties.snapToGrid = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_SNAPTOGRID) == LVS_EX_SNAPTOGRID);
	}

	*pValue = BOOL2VARIANTBOOL(properties.snapToGrid);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_SnapToGrid(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SNAPTOGRID);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.snapToGrid != b) {
		properties.snapToGrid = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			if(properties.snapToGrid) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SNAPTOGRID, LVS_EX_SNAPTOGRID);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SNAPTOGRID, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_SNAPTOGRID);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_SortOrder(SortOrderConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SortOrderConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.sortOrder;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_SortOrder(SortOrderConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SORTORDER);
	if(properties.sortOrder != newValue) {
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_ChangingSortOrder(properties.sortOrder, newValue, &cancel);
		if(cancel == VARIANT_FALSE) {
			SortOrderConstants buffer = properties.sortOrder;
			properties.sortOrder = newValue;
			SetDirty(TRUE);
			FireOnChanged(DISPID_EXLVW_SORTORDER);
			Raise_ChangedSortOrder(buffer, properties.sortOrder);
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TextBackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(LVM_GETTEXTBKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.textBackColor = CLR_NONE;
		} else if(color != OLECOLOR2COLORREF(properties.textBackColor)) {
			properties.textBackColor = color;
		}
	}

	*pValue = properties.textBackColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TextBackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TEXTBACKCOLOR);
	if(properties.textBackColor != newValue) {
		properties.textBackColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(LVM_SETTEXTBKCOLOR, 0, (properties.textBackColor == CLR_NONE ? CLR_NONE : OLECOLOR2COLORREF(properties.textBackColor)));
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_TEXTBACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewItemLines(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_COLUMNS;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		properties.tileViewItemLines = tileViewInfo.cLines;
	}

	*pValue = properties.tileViewItemLines;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewItemLines(LONG newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWITEMLINES);
	if((newValue < 0) || (newValue > 20)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.tileViewItemLines != newValue) {
		properties.tileViewItemLines = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.cLines = properties.tileViewItemLines;
			tileViewInfo.dwMask = LVTVIM_COLUMNS;
			SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWITEMLINES);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewLabelMarginBottom(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_LABELMARGIN;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		properties.tileViewLabelMarginBottom = tileViewInfo.rcLabelMargin.bottom;
	}

	*pValue = properties.tileViewLabelMarginBottom;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewLabelMarginBottom(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWLABELMARGINBOTTOM);
	if(properties.tileViewLabelMarginBottom != newValue) {
		properties.tileViewLabelMarginBottom = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.dwMask = LVTVIM_LABELMARGIN;

			SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			tileViewInfo.rcLabelMargin.bottom = properties.tileViewLabelMarginBottom;
			SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWLABELMARGINBOTTOM);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewLabelMarginLeft(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_LABELMARGIN;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		properties.tileViewLabelMarginLeft = tileViewInfo.rcLabelMargin.left;
	}

	*pValue = properties.tileViewLabelMarginLeft;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewLabelMarginLeft(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWLABELMARGINLEFT);
	if(properties.tileViewLabelMarginLeft != newValue) {
		properties.tileViewLabelMarginLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.dwMask = LVTVIM_LABELMARGIN;

			SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			tileViewInfo.rcLabelMargin.left = properties.tileViewLabelMarginLeft;
			SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWLABELMARGINLEFT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewLabelMarginRight(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_LABELMARGIN;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		properties.tileViewLabelMarginRight = tileViewInfo.rcLabelMargin.right;
	}

	*pValue = properties.tileViewLabelMarginRight;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewLabelMarginRight(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWLABELMARGINRIGHT);
	if(properties.tileViewLabelMarginRight != newValue) {
		properties.tileViewLabelMarginRight = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.dwMask = LVTVIM_LABELMARGIN;

			SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			tileViewInfo.rcLabelMargin.right = properties.tileViewLabelMarginRight;
			SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWLABELMARGINRIGHT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewLabelMarginTop(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_LABELMARGIN;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		properties.tileViewLabelMarginTop = tileViewInfo.rcLabelMargin.top;
	}

	*pValue = properties.tileViewLabelMarginTop;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewLabelMarginTop(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWLABELMARGINTOP);
	if(properties.tileViewLabelMarginTop != newValue) {
		properties.tileViewLabelMarginTop = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.dwMask = LVTVIM_LABELMARGIN;

			SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			tileViewInfo.rcLabelMargin.top = properties.tileViewLabelMarginTop;
			SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWLABELMARGINTOP);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewSubItemForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		CComPtr<IVisualProperties> pVisualProperties = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IVisualProperties), reinterpret_cast<LPARAM>(&pVisualProperties))) {
			ATLASSUME(pVisualProperties);
			COLORREF color;
			ATLVERIFY(SUCCEEDED(pVisualProperties->GetColor(VPCF_SUBTEXT, &color)));
			if(color == properties.defaultTileViewSubItemForeColor) {
				properties.tileViewSubItemForeColor = static_cast<OLE_COLOR>(-1);
			} else if(color != OLECOLOR2COLORREF(properties.tileViewSubItemForeColor)) {
				properties.tileViewSubItemForeColor = color;
			}
		}
	}

	*pValue = properties.tileViewSubItemForeColor;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewSubItemForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWSUBITEMFORECOLOR);
	if(properties.tileViewSubItemForeColor != newValue) {
		properties.tileViewSubItemForeColor = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			CComPtr<IVisualProperties> pVisualProperties = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IVisualProperties), reinterpret_cast<LPARAM>(&pVisualProperties))) {
				ATLASSUME(pVisualProperties);
				ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SUBTEXT, (properties.tileViewSubItemForeColor == -1 ? properties.defaultTileViewSubItemForeColor : OLECOLOR2COLORREF(properties.tileViewSubItemForeColor)))));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWSUBITEMFORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewTileHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_TILESIZE;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		if(tileViewInfo.dwFlags & LVTVIF_FIXEDHEIGHT) {
			properties.tileViewTileHeight = tileViewInfo.sizeTile.cy;
		} else {
			properties.tileViewTileHeight = -1;
		}
	}

	*pValue = properties.tileViewTileHeight;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewTileHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWTILEHEIGHT);
	if(properties.tileViewTileHeight != newValue) {
		properties.tileViewTileHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.dwMask = LVTVIM_TILESIZE;
			SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			if(properties.tileViewTileHeight == -1) {
				tileViewInfo.dwFlags &= ~LVTVIF_FIXEDHEIGHT;
			} else {
				tileViewInfo.dwFlags |= LVTVIF_FIXEDHEIGHT;
				tileViewInfo.sizeTile.cy = properties.tileViewTileHeight;
			}
			SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWTILEHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_TileViewTileWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_TILESIZE;
		SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
		if(tileViewInfo.dwFlags & LVTVIF_FIXEDWIDTH) {
			properties.tileViewTileWidth = tileViewInfo.sizeTile.cx;
		} else {
			properties.tileViewTileWidth = -1;
		}
	}

	*pValue = properties.tileViewTileWidth;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_TileViewTileWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TILEVIEWTILEWIDTH);
	if(properties.tileViewTileWidth != newValue) {
		properties.tileViewTileWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			LVTILEVIEWINFO tileViewInfo = {0};
			tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
			tileViewInfo.dwMask = LVTVIM_TILESIZE;

			SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			if(properties.tileViewTileWidth == -1) {
				tileViewInfo.dwFlags &= ~LVTVIF_FIXEDWIDTH;
			} else {
				tileViewInfo.dwFlags |= LVTVIF_FIXEDWIDTH;
				tileViewInfo.sizeTile.cx = properties.tileViewTileWidth;
			}

			if(IsInViewMode(vTiles, FALSE)) {
				// force an update to the item rectangles (necessary for width changes only)
				UINT b = properties.dontRedraw;
				if(!properties.dontRedraw) {
					properties.dontRedraw = TRUE;
					SetRedraw(FALSE);
				}
				SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
				SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
				SendMessage(LVM_SETVIEW, LV_VIEW_TILE, 0);
				if(properties.dontRedraw != b) {
					properties.dontRedraw = b;
					SetRedraw(!b);
				}
			} else {
				SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
			}
		}
		FireOnChanged(DISPID_EXLVW_TILEVIEWTILEWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_ToolTips(ToolTipsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ToolTipsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0));
		properties.toolTips = static_cast<ToolTipsConstants>(0);
		if((style & LVS_EX_LABELTIP)/* && IsComctl32Version580OrNewer()*/) {
			properties.toolTips = static_cast<ToolTipsConstants>(static_cast<long>(properties.toolTips) | ttLabelTipsAlways);;
		}
		if(style & LVS_EX_INFOTIP) {
			properties.toolTips = static_cast<ToolTipsConstants>(static_cast<long>(properties.toolTips) | ttInfoTips);;
		}
	}

	*pValue = properties.toolTips;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_ToolTips(ToolTipsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_TOOLTIPS);
	if(properties.toolTips != newValue) {
		properties.toolTips = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if((properties.toolTips & ttLabelTipsAlways)/* && IsComctl32Version580OrNewer()*/) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, LVS_EX_LABELTIP);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, 0);
			}
			if(properties.toolTips & ttInfoTips) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_INFOTIP, LVS_EX_INFOTIP);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_INFOTIP, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_TOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_UnderlinedItems(UnderlinedItemsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, UnderlinedItemsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0));
		properties.underlinedItems = static_cast<UnderlinedItemsConstants>(0);
		if(style & LVS_EX_UNDERLINEHOT) {
			properties.underlinedItems = static_cast<UnderlinedItemsConstants>(static_cast<long>(properties.underlinedItems) | uiHot);;
		}
		if(style & LVS_EX_UNDERLINECOLD) {
			properties.underlinedItems = static_cast<UnderlinedItemsConstants>(static_cast<long>(properties.underlinedItems) | uiCold);;
		}
	}

	*pValue = properties.underlinedItems;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_UnderlinedItems(UnderlinedItemsConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_UNDERLINEDITEMS);
	if(properties.underlinedItems != newValue) {
		properties.underlinedItems = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.underlinedItems & uiHot) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINEHOT, LVS_EX_UNDERLINEHOT);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINEHOT, 0);
			}
			if(properties.underlinedItems & uiCold) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINECOLD, LVS_EX_UNDERLINECOLD);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINECOLD, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_UNDERLINEDITEMS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_UseMinColumnWidths(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.useMinColumnWidths = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_COLUMNSNAPPOINTS) == LVS_EX_COLUMNSNAPPOINTS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useMinColumnWidths);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_UseMinColumnWidths(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_USEMINCOLUMNWIDTHS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useMinColumnWidths != b) {
		properties.useMinColumnWidths = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.useMinColumnWidths) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_COLUMNSNAPPOINTS, LVS_EX_COLUMNSNAPPOINTS);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_COLUMNSNAPPOINTS, 0);
			}
		}
		FireOnChanged(DISPID_EXLVW_USEMINCOLUMNWIDTHS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_EXLVW_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_UseWorkAreas(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useWorkAreas = ((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_MULTIWORKAREAS) == LVS_EX_MULTIWORKAREAS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useWorkAreas);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_UseWorkAreas(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_USEWORKAREAS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useWorkAreas != b) {
		properties.useWorkAreas = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.useWorkAreas) {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_MULTIWORKAREAS, LVS_EX_MULTIWORKAREAS);
			} else {
				SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_MULTIWORKAREAS, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_EXLVW_USEWORKAREAS);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_View(ViewConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ViewConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(RunTimeHelper::IsCommCtrl6()) {
			switch(SendMessage(LVM_GETVIEW, 0, 0)) {
				case LV_VIEW_ICON:
					properties.view = vIcons;
					break;
				case LV_VIEW_SMALLICON:
					properties.view = vSmallIcons;
					break;
				case LV_VIEW_LIST:
					properties.view = vList;
					break;
				case LV_VIEW_DETAILS:
					properties.view = vDetails;
					break;
				case LV_VIEW_TILE:
					properties.view = vTiles;
					if(IsComctl32Version610OrNewer()) {
						LVTILEVIEWINFO tileViewInfo = {0};
						tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
						tileViewInfo.dwMask = LVTVIM_TILESIZE;
						SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
						if(tileViewInfo.dwFlags & LVTVIF_EXTENDED) {
							properties.view = vExtendedTiles;
						}
					}
					break;
			}
		} else {
			switch(GetStyle() & LVS_TYPEMASK) {
				case LVS_ICON:
					properties.view = vIcons;
					break;
				case LVS_SMALLICON:
					properties.view = vSmallIcons;
					break;
				case LVS_LIST:
					properties.view = vList;
					break;
				case LVS_REPORT:
					properties.view = vDetails;
					break;
			}
		}
	}

	*pValue = properties.view;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_View(ViewConstants newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_VIEW);
	if(properties.view != newValue) {
		properties.view = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(RunTimeHelper::IsCommCtrl6()) {
				switch(properties.view) {
					case vIcons:
						SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
						break;
					case vSmallIcons:
						SendMessage(LVM_SETVIEW, LV_VIEW_SMALLICON, 0);
						break;
					case vList:
						SendMessage(LVM_SETVIEW, LV_VIEW_LIST, 0);
						break;
					case vDetails:
						SendMessage(LVM_SETVIEW, LV_VIEW_DETAILS, 0);
						break;
					case vTiles: {
						BOOL useHack = FALSE;
						UINT b = FALSE;
						if(IsComctl32Version610OrNewer()) {
							LVTILEVIEWINFO tileViewInfo = {0};
							tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
							tileViewInfo.dwMask = LVTVIM_TILESIZE;
							SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
							// switching directly from extended tile view to normal tile view leads to messed up items
							useHack = ((tileViewInfo.dwFlags & LVTVIF_EXTENDED) == LVTVIF_EXTENDED);
							tileViewInfo.dwFlags &= ~LVTVIF_EXTENDED;
							if(useHack) {
								// force an update to the item rectangles (necessary for width changes only)
								b = properties.dontRedraw;
								if(!properties.dontRedraw) {
									properties.dontRedraw = TRUE;
									SetRedraw(FALSE);
								}
								SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
							}
							SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
						}
						SendMessage(LVM_SETVIEW, LV_VIEW_TILE, 0);
						if(useHack) {
							if(properties.dontRedraw != b) {
								properties.dontRedraw = b;
								SetRedraw(!b);
							}
						}
						break;
					}
					case vExtendedTiles:
						if(IsComctl32Version610OrNewer()) {
							LVTILEVIEWINFO tileViewInfo = {0};
							tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
							tileViewInfo.dwMask = LVTVIM_TILESIZE;
							SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
							tileViewInfo.dwFlags |= LVTVIF_EXTENDED;
							SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
						}
						SendMessage(LVM_SETVIEW, LV_VIEW_TILE, 0);
						break;
				}
			} else {
				#ifdef INCLUDESHELLBROWSERINTERFACE
					WPARAM viewMode = 0;
				#endif
				switch(properties.view) {
					case vIcons:
						ModifyStyle(LVS_SMALLICON | LVS_LIST | LVS_REPORT, LVS_ICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						#ifdef INCLUDESHELLBROWSERINTERFACE
							viewMode = LV_VIEW_ICON;
						#endif
						break;
					case vSmallIcons:
						ModifyStyle(LVS_ICON | LVS_LIST | LVS_REPORT, LVS_SMALLICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						#ifdef INCLUDESHELLBROWSERINTERFACE
							viewMode = LV_VIEW_SMALLICON;
						#endif
						break;
					case vList:
						ModifyStyle(LVS_ICON | LVS_SMALLICON | LVS_REPORT, LVS_LIST, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						#ifdef INCLUDESHELLBROWSERINTERFACE
							viewMode = LV_VIEW_LIST;
						#endif
						break;
					case vDetails:
						ModifyStyle(LVS_ICON | LVS_SMALLICON | LVS_LIST, LVS_REPORT, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						#ifdef INCLUDESHELLBROWSERINTERFACE
							viewMode = LV_VIEW_DETAILS;
						#endif
						break;
				}
				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(shellBrowserInterface.pInternalMessageListener) {
						shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_CHANGEDVIEW, viewMode, 0);
					}
				#endif
			}
		}
		FireOnChanged(DISPID_EXLVW_VIEW);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_VirtualItemCount(VARIANT_BOOL/* noScroll = VARIANT_FALSE*/, VARIANT_BOOL/* noInvalidateAll = VARIANT_TRUE*/, LONG* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.virtualItemCount;
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_VirtualItemCount(VARIANT_BOOL noScroll/* = VARIANT_FALSE*/, VARIANT_BOOL noInvalidateAll/* = VARIANT_TRUE*/, LONG newValue/* = 0*/)
{
	PUTPROPPROLOG(DISPID_EXLVW_VIRTUALITEMCOUNT);
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.virtualItemCount != newValue) {
		properties.virtualItemCount = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD flags = 0;
			if(noScroll != VARIANT_FALSE) {
				flags |= LVSICF_NOSCROLL;
			}
			if(noInvalidateAll != VARIANT_FALSE) {
				flags |= LVSICF_NOINVALIDATEALL;
			}
			SendMessage(LVM_SETITEMCOUNT, properties.virtualItemCount, flags);
		}
		FireOnChanged(DISPID_EXLVW_VIRTUALITEMCOUNT);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_VirtualMode(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.virtualMode = ((GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA);
	}

	*pValue = BOOL2VARIANTBOOL(properties.virtualMode);
	return S_OK;
}

STDMETHODIMP ExplorerListView::put_VirtualMode(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_EXLVW_VIRTUALMODE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.virtualMode != b) {
		properties.virtualMode = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_EXLVW_VIRTUALMODE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::get_WorkAreas(IListViewWorkAreas** ppWorkAreas)
{
	ATLASSERT_POINTER(ppWorkAreas, IListViewWorkAreas*);
	if(!ppWorkAreas) {
		return E_POINTER;
	}

	ClassFactory::InitWorkAreas(this, IID_IListViewWorkAreas, reinterpret_cast<LPUNKNOWN*>(ppWorkAreas));
	return S_OK;
}

STDMETHODIMP ExplorerListView::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ExplorerListView::ApproximateViewRectangle(LONG numberOfItems/* = -1*/, OLE_XSIZE_PIXELS* pProposedWidth/* = NULL*/, OLE_YSIZE_PIXELS* pProposedHeight/* = NULL*/)
{
	if(IsWindow()) {
		int cx = -1;
		if(pProposedWidth) {
			cx = *pProposedWidth;
		}
		int cy = -1;
		if(pProposedHeight) {
			cy = *pProposedHeight;
		}
		DWORD approximatedSize = static_cast<DWORD>(SendMessage(LVM_APPROXIMATEVIEWRECT, numberOfItems, MAKELPARAM(cx, cy)));
		if(pProposedWidth) {
			*pProposedWidth = WTL::CSize(approximatedSize).cx;
		}
		if(pProposedHeight) {
			*pProposedHeight = WTL::CSize(approximatedSize).cy;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::ArrangeItems(ArrangementStyleConstants arrangementStyle/* = astDefault*/)
{
	if(IsWindow()) {
		if(SendMessage(LVM_ARRANGE, arrangementStyle, 0)) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::BeginDrag(IListViewItemContainer* pDraggedItems, OLE_HANDLE hDragImageList/* = NULL*/, OLE_XPOS_PIXELS* pXHotSpot/* = NULL*/, OLE_YPOS_PIXELS* pYHotSpot/* = NULL*/)
{
	ATLASSUME(pDraggedItems);
	if(!pDraggedItems) {
		return E_POINTER;
	}

	int xHotSpot = 0;
	if(pXHotSpot) {
		xHotSpot = *pXHotSpot;
	}
	int yHotSpot = 0;
	if(pYHotSpot) {
		yHotSpot = *pYHotSpot;
	}
	HRESULT hr = dragDropStatus.BeginDrag(*this, pDraggedItems, static_cast<HIMAGELIST>(LongToHandle(hDragImageList)), &xHotSpot, &yHotSpot);
	SetCapture();
	if(pXHotSpot) {
		*pXHotSpot = xHotSpot;
	}
	if(pYHotSpot) {
		*pYHotSpot = yHotSpot;
	}

	if(dragDropStatus.hDragImageList) {
		ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
		dragDropStatus.dragImageIsHidden = 0;
		ImageList_DragEnter(0, 0, 0);

		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}
	return hr;
}

STDMETHODIMP ExplorerListView::CountVisible(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(IsInViewMode(vList) || IsInViewMode(vDetails)) {
			*pValue = static_cast<LONG>(SendMessage(LVM_GETCOUNTPERPAGE, 0, 0));
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::CreateItemContainer(VARIANT items/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListViewItemContainer** ppContainer/* = NULL*/)
{
	ATLASSERT_POINTER(ppContainer, IListViewItemContainer*);
	if(!ppContainer) {
		return E_POINTER;
	}

	*ppContainer = NULL;
	CComObject<ListViewItemContainer>* pLvwItemContainerObj = NULL;
	CComObject<ListViewItemContainer>::CreateInstance(&pLvwItemContainerObj);
	pLvwItemContainerObj->AddRef();

	// clone all settings
	pLvwItemContainerObj->SetOwner(this);
	pLvwItemContainerObj->properties.useIndexes = ((GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA);

	pLvwItemContainerObj->QueryInterface(__uuidof(IListViewItemContainer), reinterpret_cast<LPVOID*>(ppContainer));
	pLvwItemContainerObj->Release();

	if(*ppContainer) {
		(*ppContainer)->Add(items);
		RegisterItemContainer(static_cast<IItemContainer*>(pLvwItemContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::EndDrag(VARIANT_BOOL abort)
{
	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	KillTimer(timers.ID_DRAGSCROLL);
	ReleaseCapture();
	if(dragDropStatus.hDragImageList) {
		dragDropStatus.HideDragImage(TRUE);
		ImageList_EndDrag();
	}

	HRESULT hr = S_OK;
	if(abort) {
		hr = Raise_AbortedDrag();
	} else {
		hr = Raise_Drop();
	}

	dragDropStatus.EndDrag();
	Invalidate();

	return hr;
}

STDMETHODIMP ExplorerListView::EndLabelEdit(VARIANT_BOOL discard)
{
	if(IsWindow()) {
		if(discard) {
			if(RunTimeHelper::IsCommCtrl6()) {
				SendMessage(LVM_CANCELEDITLABEL, 0, 0);
			} else {
				SendMessage(LVM_EDITLABEL, static_cast<WPARAM>(-1), 0);
			}
		} else {
			if(containedEdit.IsWindow()) {
				containedEdit.SendMessage(WM_KEYDOWN, VK_RETURN, 0);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::FindItem(SearchModeConstants searchMode, VARIANT searchFor, SearchDirectionConstants searchDirection/* = sdNoneSpecific*/, VARIANT_BOOL wrapAtLastItem/* = VARIANT_TRUE*/, IListViewItem** ppFoundItem/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		CComPtr<IListView_WIN7> pListView7 = NULL;
		CComPtr<IListView_WINVISTA> pListViewVista = NULL;
		if(!SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7))) {
			SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista));
		}
		if(pListView7 || pListViewVista) {
			LVFINDINFOW findInfo = {0};
			VARIANT v;
			VariantInit(&v);
			switch(searchMode) {
				case smItemData:
					findInfo.flags = LVFI_PARAM;
					if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_UI4))) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					findInfo.lParam = v.ulVal;
					break;
				case smPartialText:
					findInfo.flags = LVFI_PARTIAL;  // could also be LVFI_SUBSTRING
					// fall through
				case smText:
					findInfo.flags |= LVFI_STRING;
					if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_BSTR))) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					findInfo.psz = OLE2W(v.bstrVal);
					break;
				case smNearestPosition:
					findInfo.flags = LVFI_NEARESTXY;
					findInfo.vkDirection = static_cast<UINT>(searchDirection);
					if((searchFor.vt & VT_ARRAY) == VT_ARRAY && searchFor.parray) {
						LONG l = 0;
						SafeArrayGetLBound(searchFor.parray, 1, &l);
						LONG u = 0;
						SafeArrayGetUBound(searchFor.parray, 1, &u);
						if(u < l || u - l + 1 != 2) {
							// invalid arg - raise VB runtime error 380
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}

						VARTYPE vt = 0;
						SafeArrayGetVartype(searchFor.parray, &vt);
						ULONG element = 0;
						for(LONG i = l; i <= u; ++i) {
							if(vt == VT_VARIANT) {
								SafeArrayGetElement(searchFor.parray, &i, &v);
								if(FAILED(VariantChangeType(&v, &v, 0, VT_UI4))) {
									// invalid arg - raise VB runtime error 380
									return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
								}
								element = v.ulVal;
							} else {
								SafeArrayGetElement(searchFor.parray, &i, &element);
							}
							if(i == l) {
								findInfo.pt.x = element;
							} else {
								findInfo.pt.y = element;
							}
						}
					} else {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
			}
			VariantClear(&v);
			if(wrapAtLastItem != VARIANT_FALSE) {
				findInfo.flags |= LVFI_WRAP;
			}

			LVITEMINDEX findItem = {-1, 0};
			LVITEMINDEX foundItem = {-1, 0};
			if(pListView7) {
				hr = pListView7->FindItem(findItem, &findInfo, &foundItem);
			} else {
				hr = pListViewVista->FindItem(findItem, &findInfo, &foundItem);
			}
			if(SUCCEEDED(hr)) {
				ClassFactory::InitListItem(foundItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
			} else {
				*ppFoundItem = NULL;
			}
			hr = S_OK;
		}
		if(FAILED(hr)) {
			LVFINDINFO findInfo = {0};
			VARIANT v;
			VariantInit(&v);
			switch(searchMode) {
				case smItemData:
					findInfo.flags = LVFI_PARAM;
					if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_UI4))) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					findInfo.lParam = v.ulVal;
					break;
				case smPartialText:
					findInfo.flags = LVFI_PARTIAL;
					// fall through
				case smText:
					findInfo.flags |= LVFI_STRING;
					if(FAILED(VariantChangeType(&v, &searchFor, 0, VT_BSTR))) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
				case smNearestPosition:
					findInfo.flags = LVFI_NEARESTXY;
					findInfo.vkDirection = static_cast<UINT>(searchDirection);
					if((searchFor.vt & VT_ARRAY) == VT_ARRAY && searchFor.parray) {
						LONG l = 0;
						SafeArrayGetLBound(searchFor.parray, 1, &l);
						LONG u = 0;
						SafeArrayGetUBound(searchFor.parray, 1, &u);
						if(u < l || u - l + 1 != 2) {
							// invalid arg - raise VB runtime error 380
							return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
						}

						VARTYPE vt = 0;
						SafeArrayGetVartype(searchFor.parray, &vt);
						ULONG element = 0;
						for(LONG i = l; i <= u; ++i) {
							if(vt == VT_VARIANT) {
								SafeArrayGetElement(searchFor.parray, &i, &v);
								if(FAILED(VariantChangeType(&v, &v, 0, VT_UI4))) {
									// invalid arg - raise VB runtime error 380
									return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
								}
								element = v.ulVal;
							} else {
								SafeArrayGetElement(searchFor.parray, &i, &element);
							}
							if(i == l) {
								findInfo.pt.x = element;
							} else {
								findInfo.pt.y = element;
							}
						}
					} else {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					break;
			}
			COLE2T converter(v.bstrVal);
			if(findInfo.flags & LVFI_STRING) {
				findInfo.psz = converter;
			}
			if(wrapAtLastItem != VARIANT_FALSE) {
				findInfo.flags |= LVFI_WRAP;
			}

			LVITEMINDEX foundItem;
			foundItem.iItem = static_cast<int>(SendMessage(LVM_FINDITEM, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&findInfo)));
			foundItem.iGroup = 0;
			ClassFactory::InitListItem(foundItem, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppFoundItem));
			VariantClear(&v);
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ExplorerListView::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP ExplorerListView::GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IListViewItem** ppListItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppListItem, IListViewItem*);
	if(!ppListItem) {
		return E_POINTER;
	}

	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		POINT pt = {x, y};
		LVINSERTMARK insertMark = {0};
		insertMark.cbSize = sizeof(LVINSERTMARK);
		if(SendMessage(LVM_INSERTMARKHITTEST, reinterpret_cast<WPARAM>(&pt), reinterpret_cast<LPARAM>(&insertMark))) {
			if(insertMark.iItem == -1) {
				*ppListItem = NULL;
				*pRelativePosition = impNowhere;
			} else {
				if(insertMark.dwFlags & LVIM_AFTER) {
					*pRelativePosition = impAfter;
				} else {
					*pRelativePosition = impBefore;
				}
				LVITEMINDEX itemIndex = {insertMark.iItem, 0};
				ClassFactory::InitListItem(itemIndex, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppListItem), FALSE);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetFooterRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		RECT footerRectangle = {0};
		if(SendMessage(LVM_GETFOOTERRECT, 0, reinterpret_cast<LPARAM>(&footerRectangle))) {
			if(pXLeft) {
				*pXLeft = footerRectangle.left;
			}
			if(pYTop) {
				*pYTop = footerRectangle.top;
			}
			if(pXRight) {
				*pXRight = footerRectangle.right;
			}
			if(pYBottom) {
				*pYBottom = footerRectangle.bottom;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetHeaderChevronRectangle(OLE_XPOS_PIXELS* pLeft/* = NULL*/, OLE_YPOS_PIXELS* pTop/* = NULL*/, OLE_XPOS_PIXELS* pRight/* = NULL*/, OLE_YPOS_PIXELS* pBottom/* = NULL*/)
{
	if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		RECT chevronRectangle = {0};
		if(containedSysHeader32.SendMessage(HDM_GETOVERFLOWRECT, 0, reinterpret_cast<LPARAM>(&chevronRectangle))) {
			if(pLeft) {
				*pLeft = chevronRectangle.left;
			}
			if(pTop) {
				*pTop = chevronRectangle.top;
			}
			if(pRight) {
				*pRight = chevronRectangle.right;
			}
			if(pBottom) {
				*pBottom = chevronRectangle.bottom;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetIconSpacing(ViewConstants viewMode, OLE_XSIZE_PIXELS* pItemWidth/* = NULL*/, OLE_YSIZE_PIXELS* pItemHeight/* = NULL*/)
{
	if(IsWindow()) {
		DWORD spacing = 0;
		switch(viewMode) {
			case vIcons:
				spacing = static_cast<DWORD>(SendMessage(LVM_GETITEMSPACING, FALSE, 0));
				break;
			case vSmallIcons:
				spacing = static_cast<DWORD>(SendMessage(LVM_GETITEMSPACING, TRUE, 0));
				break;
			default:
				return E_FAIL;
				break;
		}
		if(pItemWidth) {
			*pItemWidth = LOWORD(spacing);
		}
		if(pItemHeight) {
			*pItemHeight = HIWORD(spacing);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IListViewItem** ppListItem)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppListItem, IListViewItem*);
	if(!ppListItem) {
		return E_POINTER;
	}

	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVINSERTMARK insertMark = {0};
		insertMark.cbSize = sizeof(LVINSERTMARK);
		if(SendMessage(LVM_GETINSERTMARK, 0, reinterpret_cast<LPARAM>(&insertMark))) {
			if(insertMark.iItem == -1) {
				*ppListItem = NULL;
				*pRelativePosition = impNowhere;
			} else {
				if(insertMark.dwFlags & LVIM_AFTER) {
					*pRelativePosition = impAfter;
				} else {
					*pRelativePosition = impBefore;
				}
				LVITEMINDEX itemIndex = {insertMark.iItem, 0};
				ClassFactory::InitListItem(itemIndex, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppListItem), FALSE);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetInsertMarkRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		RECT rc = {0};
		if(SendMessage(LVM_GETINSERTMARKRECT, 0, reinterpret_cast<LPARAM>(&rc))) {
			if(pXLeft) {
				*pXLeft = rc.left;
			}
			if(pYTop) {
				*pYTop = rc.top;
			}
			if(pXRight) {
				*pXRight = rc.right;
			}
			if(pYBottom) {
				*pYBottom = rc.bottom;
			}
			return S_OK;
		} else {
			// there's no insertion mark
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetOrigin(OLE_XPOS_PIXELS* pX/* = NULL*/, OLE_YPOS_PIXELS* pY/* = NULL*/)
{
	if(IsWindow()) {
		POINT pt = {0};
		if(SendMessage(LVM_GETORIGIN, 0, reinterpret_cast<LPARAM>(&pt))) {
			if(pX) {
				*pX = pt.x;
			}
			if(pY) {
				*pY = pt.y;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetStringWidth(BSTR stringToTest, OLE_XSIZE_PIXELS* pWidth)
{
	ATLASSERT_POINTER(pWidth, OLE_XSIZE_PIXELS);
	if(!pWidth) {
		return E_POINTER;
	}

	if(IsWindow()) {
		LPCTSTR pStringToTest = COLE2CT(stringToTest);
		*pWidth = static_cast<OLE_XSIZE_PIXELS>(SendMessage(LVM_GETSTRINGWIDTH, 0, reinterpret_cast<LPARAM>(pStringToTest)));
		if(*pWidth != 0) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::GetViewRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(IsWindow()) {
		if(!IsInViewMode(vList) && !IsInViewMode(vDetails)) {
			RECT rc = {0};
			if(SendMessage(LVM_GETVIEWRECT, 0, reinterpret_cast<LPARAM>(&rc))) {
				if(pXLeft) {
					*pXLeft = rc.left;
				}
				if(pYTop) {
					*pYTop = rc.top;
				}
				if(pXRight) {
					*pXRight = rc.right;
				}
				if(pYBottom) {
					*pYBottom = rc.bottom;
				}
				return S_OK;
			}
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::HeaderBeginDrag(IListViewColumn* pDraggedColumn, OLE_HANDLE hDragImageList/* = NULL*/, OLE_XPOS_PIXELS* pXHotSpot/* = NULL*/, OLE_YPOS_PIXELS* pYHotSpot/* = NULL*/, VARIANT_BOOL restrictDragImage/* = VARIANT_TRUE*/)
{
	ATLASSUME(pDraggedColumn);
	if(!pDraggedColumn) {
		return E_POINTER;
	}

	int xHotSpot = 0;
	if(pXHotSpot) {
		xHotSpot = *pXHotSpot;
	}
	int yHotSpot = 0;
	if(pYHotSpot) {
		yHotSpot = *pYHotSpot;
	}
	HRESULT hr = dragDropStatus.HeaderBeginDrag(containedSysHeader32.m_hWnd, pDraggedColumn, static_cast<HIMAGELIST>(LongToHandle(hDragImageList)), &xHotSpot, &yHotSpot);
	if(pXHotSpot) {
		*pXHotSpot = xHotSpot;
	}
	if(pYHotSpot) {
		*pYHotSpot = yHotSpot;
	}
	// avoid any inapplicable Header*Click events
	mouseStatus_Header.RemoveAllClickCandidates();
	containedSysHeader32.SetCapture();

	if(dragDropStatus.hDragImageList) {
		dragDropStatus.restrictedDragImage = VARIANTBOOL2BOOL(restrictDragImage);
		if(dragDropStatus.restrictedDragImage) {
			yHotSpot = 0;
			if(pYHotSpot) {
				*pYHotSpot = yHotSpot;
			}
			ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
			dragDropStatus.dragImageIsHidden = 0;
			ImageList_DragEnter(0, 0, 0);

			RECT windowRect = {0};
			containedSysHeader32.GetWindowRect(&windowRect);
			ImageList_DragMove(GET_X_LPARAM(GetMessagePos()), windowRect.top);
		} else {
			ImageList_BeginDrag(dragDropStatus.hDragImageList, 0, xHotSpot, yHotSpot);
			dragDropStatus.dragImageIsHidden = 0;
			ImageList_DragEnter(0, 0, 0);

			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ImageList_DragMove(mousePosition.x, mousePosition.y);
		}
	}
	return hr;
}

STDMETHODIMP ExplorerListView::HeaderEndDrag(VARIANT_BOOL abort)
{
	if(!dragDropStatus.HeaderIsDragging()) {
		return E_FAIL;
	}

	KillTimer(timers.ID_DRAGSCROLL);
	ReleaseCapture();
	if(dragDropStatus.hDragImageList) {
		dragDropStatus.HideDragImage(TRUE);
		ImageList_EndDrag();
	}

	HRESULT hr = S_OK;
	if(abort) {
		hr = Raise_HeaderAbortedDrag();
	} else {
		hr = Raise_HeaderDrop();
	}

	dragDropStatus.HeaderEndDrag();

	return hr;
}

STDMETHODIMP ExplorerListView::HeaderHitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants* pHitTestDetails, IListViewColumn** ppHitColumn)
{
	ATLASSERT_POINTER(ppHitColumn, IListViewColumn*);
	if(!ppHitColumn) {
		return E_POINTER;
	}

	if(containedSysHeader32.IsWindow()) {
		UINT flags = static_cast<UINT>(*pHitTestDetails);
		int column = HeaderHitTest(x, y, &flags);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HeaderHitTestConstants>(flags);
		}
		ClassFactory::InitColumn(column, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(ppHitColumn));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::HeaderOLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, IListViewColumn* pDraggedColumnHeader/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSUME(pDraggedColumnHeader);
		if(!pDraggedColumnHeader) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = containedSysHeader32;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay == -1) {
			if(pDraggedColumnHeader) {
				if(flags.usingThemes && RunTimeHelper::IsVista()) {
					pOLEDataObjectObj->properties.numberOfItemsToDisplay = 1;
				}
			}
		} else if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedColumnHeader) {
		LONG l = -1;
		pDraggedColumnHeader->get_Index(&l);
		dragDropStatus.draggedColumn = l;
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	containedSysHeader32.ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is ExplorerListView[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.headerOLEDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	dragDropStatus.headerIsSource = TRUE;
	if(pOLEDataObject) {
		Raise_HeaderOLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_HeaderOLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}
	dragDropStatus.headerIsSource = FALSE;

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP ExplorerListView::HeaderRefresh(void)
{
	if(containedSysHeader32.IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		containedSysHeader32.Invalidate();
		containedSysHeader32.UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, VARIANT* pHitSubItem/* = NULL*/, VARIANT* pHitGroup/* = NULL*/, VARIANT* pHitFooterItem/* = NULL*/, IListViewItem** ppHitItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppHitItem, IListViewItem*);
	if(!ppHitItem) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT flags = HitTestConstants2LVHTFlags(*pHitTestDetails);
		LVITEMINDEX itemIndex = {-1, 0};
		if(pHitSubItem && pHitSubItem->vt != VT_ERROR) {
			int subItemIndex = -1;
			HitTest(x, y, &flags, &itemIndex, &subItemIndex, TRUE);
			VariantClear(pHitSubItem);
			ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(pHitSubItem->ppdispVal));
			pHitSubItem->vt = VT_DISPATCH | VT_BYREF;
		} else {
			HitTest(x, y, &flags, &itemIndex, NULL, TRUE);
		}

		if(IsComctl32Version610OrNewer()) {
			if(pHitGroup && pHitGroup->vt != VT_ERROR) {
				VariantClear(pHitGroup);
				if(flags & LVHT_EX_GROUP) {
					itemIndex.iItem = -1;
					itemIndex.iGroup = 0;
					HitTest(x, y, NULL, &itemIndex, NULL, TRUE, TRUE, FALSE);
					ClassFactory::InitGroup(itemIndex.iItem, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(pHitGroup->ppdispVal));
					pHitGroup->vt = VT_DISPATCH | VT_BYREF;
				}
			}
			if(pHitFooterItem && pHitFooterItem->vt != VT_ERROR) {
				VariantClear(pHitFooterItem);
				if(flags & LVHT_EX_FOOTER) {
					itemIndex.iItem = -1;
					itemIndex.iGroup = 0;
					HitTest(x, y, NULL, &itemIndex, NULL, TRUE, TRUE, FALSE);
					ClassFactory::InitFooterItem(itemIndex.iItem, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(pHitFooterItem->ppdispVal));
					pHitFooterItem->vt = VT_DISPATCH | VT_BYREF;
				}
			}
		}

		if(pHitTestDetails) {
			*pHitTestDetails = LVHTFlags2HitTestConstants(flags, itemIndex.iItem != -1);
		}
		ClassFactory::InitListItem(itemIndex, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppHitItem));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::IsFooterVisible(VARIANT_BOOL* pVisible)
{
	HRESULT hr = E_FAIL;
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		CComPtr<IListViewFooter> pListViewFooter = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListViewFooter), reinterpret_cast<LPARAM>(&pListViewFooter))) {
			ATLASSUME(pListViewFooter);
			int visible = FALSE;
			hr = pListViewFooter->IsVisible(&visible);
			*pVisible = BOOL2VARIANTBOOL(visible);
		}
	}
	return hr;
}

STDMETHODIMP ExplorerListView::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, IListViewItemContainer* pDraggedItems/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSUME(pDraggedItems);
		if(!pDraggedItems) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = *this;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay == -1) {
			if(pDraggedItems) {
				if(flags.usingThemes && RunTimeHelper::IsVista()) {
					pDraggedItems->Count(&pOLEDataObjectObj->properties.numberOfItemsToDisplay);
				}
			}
		} else if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedItems) {
		pDraggedItems->Clone(&dragDropStatus.pDraggedItems);
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is ExplorerListView[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	dragDropStatus.headerIsSource = FALSE;
	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	KillTimer(timers.ID_DRAGSCROLL);
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP ExplorerListView::RedrawItems(LONG iFirstItem/* = 0*/, LONG iLastItem/* = -1*/)
{
	if(IsWindow()) {
		if(iLastItem == -1) {
			iLastItem = static_cast<int>(SendMessage(LVM_GETITEMCOUNT, 0, 0)) - 1;
		}
		if(SendMessage(LVM_REDRAWITEMS, iFirstItem, iLastItem)) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::Refresh(void)
{
	if(IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ExplorerListView::Scroll(OLE_XSIZE_PIXELS horizontalDistance, OLE_YSIZE_PIXELS verticalDistance)
{
	if(IsWindow()) {
		if(SendMessage(LVM_SCROLL, horizontalDistance, verticalDistance)) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::SetFocusToHeader(void)
{
	if(containedSysHeader32.IsWindow() && IsComctl32Version610OrNewer()) {
		HWND hWndFocus = GetFocus();
		if((hWndFocus != *this) && !IsChild(hWndFocus)) {
			// this will handle the OLE stuff
			SetFocus();
		}
		// do the real focus change
		containedSysHeader32.SetFocus();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::SetHeaderInsertMarkPosition(LONG position/* = 0*/, LONG x/* = 0x80000000*/, LONG y/* = 0x80000000*/, LONG* pPosition/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(containedSysHeader32.IsWindow()) {
		dragDropStatus.HideDragImage(TRUE);
		if((x == 0x80000000) || (y == 0x80000000)) {
			if((position >= -1) && (position <= static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)))) {
				if(containedSysHeader32.SendMessage(HDM_SETHOTDIVIDER, FALSE, position) == position) {
					if(pPosition) {
						*pPosition = position;
					}
					hr = S_OK;
				}
			}
		} else {
			position = static_cast<LONG>(containedSysHeader32.SendMessage(HDM_SETHOTDIVIDER, TRUE, MAKELPARAM(x, y)));
			if(pPosition) {
				*pPosition = position;
			}
			hr = S_OK;
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	return hr;
}

STDMETHODIMP ExplorerListView::SetIconSpacing(OLE_XSIZE_PIXELS itemWidth, OLE_YSIZE_PIXELS itemHeight)
{
	if((itemWidth < -1) || (itemWidth > 0xFFFF)) {
		return E_INVALIDARG;
	}
	if((itemHeight < -1) || (itemHeight > 0xFFFF)) {
		return E_INVALIDARG;
	}
	if((itemWidth == -1) && (itemHeight != -1)) {
		return E_INVALIDARG;
	}
	if((itemWidth != -1) && (itemHeight == -1)) {
		return E_INVALIDARG;
	}

	if(IsWindow()) {
		if((itemWidth == -1) && (itemHeight == -1)) {
			SendMessage(LVM_SETICONSPACING, 0, -1);
		} else {
			SendMessage(LVM_SETICONSPACING, 0, MAKELONG(LOWORD(itemWidth), LOWORD(itemHeight)));
		}
		Invalidate();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ExplorerListView::SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, IListViewItem* pListItem)
{
	int itemIndex = -1;
	if(pListItem) {
		LONG l = -1;
		pListItem->get_Index(&l);
		itemIndex = l;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		LVINSERTMARK insertMark = {0};
		insertMark.cbSize = sizeof(LVINSERTMARK);
		insertMark.iItem = itemIndex;
		switch(relativePosition) {
			case impNowhere:
				insertMark.iItem = -1;
				break;
			case impAfter:
				insertMark.dwFlags |= LVIM_AFTER;
				break;
			case impDontChange:
				return S_OK;
				break;
		}

		dragDropStatus.HideDragImage(TRUE);
		if(SendMessage(LVM_SETINSERTMARK, 0, reinterpret_cast<LPARAM>(&insertMark))) {
			hr = S_OK;
		}
		dragDropStatus.ShowDragImage(TRUE);
	}

	return hr;
}

STDMETHODIMP ExplorerListView::ShowFooter(void)
{
	HRESULT hr = E_FAIL;
	if(IsWindow() && IsComctl32Version610OrNewer()) {
		CComPtr<IListViewFooter> pListViewFooter = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListViewFooter), reinterpret_cast<LPARAM>(&pListViewFooter))) {
			ATLASSUME(pListViewFooter);
			ATLVERIFY(SUCCEEDED(pListViewFooter->SetIntroText(OLE2W(properties.footerIntroText))));
			CComPtr<IListViewFooterCallback> pListViewFooterCallback = NULL;
			QueryInterface(IID_IListViewFooterCallback, reinterpret_cast<LPVOID*>(&pListViewFooterCallback));
			hr = pListViewFooter->Show(pListViewFooterCallback);
		}
	}
	return hr;
}

STDMETHODIMP ExplorerListView::SortGroups(SortByConstants firstCriterion/* = sobShell*/, SortByConstants secondCriterion/* = sobText*/, SortByConstants thirdCriterion/* = sobNone*/, VARIANT_BOOL caseSensitive/* = VARIANT_FALSE*/)
{
	if(!RunTimeHelper::IsCommCtrl6()) {
		return E_FAIL;
	}

	ATLASSERT(IsWindow());

	if(firstCriterion == sobNone) {
		// nothing to do...
		return S_OK;
	}
	#ifndef INCLUDESHELLBROWSERINTERFACE
		if((firstCriterion == sobShell) && (secondCriterion == sobNone)) {
			// nothing to do...
			return S_OK;
		}
	#endif

	sortingSettings.caseSensitive = VARIANTBOOL2BOOL(caseSensitive);
	sortingSettings.sortingCriteria[0] = firstCriterion;
	sortingSettings.sortingCriteria[1] = secondCriterion;
	sortingSettings.sortingCriteria[2] = thirdCriterion;
	sortingSettings.localeID = properties.groupLocale;
	sortingSettings.flagsForVarFromStr = properties.groupTextParsingFlagsForVarFromStr;
	sortingSettings.flagsForCompareString = properties.groupTextParsingFlagsForCompareString;

	if(!SendMessage(LVM_SORTGROUPS, reinterpret_cast<WPARAM>(::CompareGroups), reinterpret_cast<LPARAM>(static_cast<IGroupComparator*>(this)))) {
		return E_FAIL;
	}

	return S_OK;
}

STDMETHODIMP ExplorerListView::SortItems(SortByConstants firstCriterion/* = sobShell*/, SortByConstants secondCriterion/* = sobText*/, SortByConstants thirdCriterion/* = sobNone*/, SortByConstants fourthCriterion/* = sobNone*/, SortByConstants fifthCriterion/* = sobNone*/, VARIANT column/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, VARIANT_BOOL caseSensitive/* = VARIANT_FALSE*/)
{
	ATLASSERT(IsWindow());

	if(firstCriterion == sobNone) {
		// nothing to do...
		return S_OK;
	}
	#ifndef INCLUDESHELLBROWSERINTERFACE
		if((firstCriterion == sobShell) && (secondCriterion == sobNone)) {
			// nothing to do...
			return S_OK;
		}
	#endif
	sortingSettings.column = 0;
	if(column.vt != VT_ERROR) {
		switch(column.vt) {
			case VT_DISPATCH:
				if(column.pdispVal) {
					CComQIPtr<IListViewColumn, &IID_IListViewColumn> pColumn(column.pdispVal);
					if(pColumn) {
						pColumn->get_Index(reinterpret_cast<LONG*>(&sortingSettings.column));
					}
				}
				break;
			default:
				VARIANT v;
				VariantInit(&v);
				if(FAILED(VariantChangeType(&v, &column, 0, VT_INT))) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				sortingSettings.column = v.intVal;
				break;
		}

		if(!IsValidColumnIndex(sortingSettings.column, this)) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	}

	sortingSettings.caseSensitive = VARIANTBOOL2BOOL(caseSensitive);
	sortingSettings.sortingCriteria[0] = firstCriterion;
	sortingSettings.sortingCriteria[1] = secondCriterion;
	sortingSettings.sortingCriteria[2] = thirdCriterion;
	sortingSettings.sortingCriteria[3] = fourthCriterion;
	sortingSettings.sortingCriteria[4] = fifthCriterion;
	ColumnData columnData = {0};
	BOOL foundColumn = FALSE;
	#ifdef USE_STL
		std::list<ColumnData>::iterator iter = columnParams.begin();
		if(iter != columnParams.end()) {
			std::advance(iter, sortingSettings.column);
			if(iter != columnParams.end()) {
				columnData = *iter;
				foundColumn = TRUE;
			}
		}
	#else
		POSITION p = columnParams.FindIndex(sortingSettings.column);
		if(p) {
			columnData = columnParams.GetAt(p);
			foundColumn = TRUE;
		}
	#endif
	if(foundColumn) {
		sortingSettings.localeID = columnData.localeID;
		sortingSettings.flagsForVarFromStr = columnData.flagsForVarFromStr;
		sortingSettings.flagsForCompareString = columnData.flagsForCompareString;
	} else {
		sortingSettings.localeID = static_cast<LCID>(-1);
		sortingSettings.flagsForVarFromStr = 0;
		sortingSettings.flagsForCompareString = 0;
	}

	//if(IsComctl32Version580OrNewer()) {
		if(!SendMessage(LVM_SORTITEMSEX, reinterpret_cast<WPARAM>(static_cast<IItemComparator*>(this)), reinterpret_cast<LPARAM>(::CompareItemsEx))) {
			return E_FAIL;
		}
	/*} else {
		if(!SendMessage(LVM_SORTITEMS, reinterpret_cast<WPARAM>(static_cast<IItemComparator*>(this)), reinterpret_cast<LPARAM>(::CompareItems))) {
			return E_FAIL;
		}
	}*/

	return S_OK;
}


LRESULT ExplorerListView::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deListKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ExplorerListView::OnContextMenu(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(reinterpret_cast<HWND>(wParam) == *this) {
		if(flags.ignoreRClickNotification) {
			return 0;
		}

		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(!flags.receivedRButtonUp) {
			if((mousePosition.x != -1) && (mousePosition.y != -1)) {
				// the message was triggered by mouse input
				return 0;
			}
		}
		flags.receivedRButtonUp = FALSE;

		if((mousePosition.x != -1) && (mousePosition.y != -1)) {
			ScreenToClient(&mousePosition);
		}

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
		return 0;
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		/* SysListView32 already sent a WM_NOTIFYFORMAT/NF_QUERY message to the parent window, but
		   unfortunately our reflection object did not yet know where to reflect this message to. So we ask
		   SysListView32 to send the message again. */
		SendMessage(WM_NOTIFYFORMAT, reinterpret_cast<WPARAM>(reinterpret_cast<LPCREATESTRUCT>(lParam)->hwndParent), NF_REQUERY);

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT ExplorerListView::OnCtlColorEdit(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.editBackColor));
	SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.editForeColor));
	return reinterpret_cast<LRESULT>(hEditBackColorBrush);
}

LRESULT ExplorerListView::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	/*ATLASSERT(pDetails->CtlType == ODT_HEADER);
	ATLASSERT(pDetails->itemAction == ODA_DRAWENTIRE);
	ATLASSERT((pDetails->itemState & (ODS_GRAYED | ODS_DISABLED | ODS_CHECKED | ODS_FOCUS | ODS_DEFAULT | ODS_COMBOBOXEDIT | ODS_INACTIVE)) == 0);*/

	if(pDetails->hwndItem == containedSysHeader32.m_hWnd) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->itemID, this, FALSE);

		Raise_HeaderOwnerDrawItem(pLvwColumn, static_cast<OwnerDrawItemStateConstants>(pDetails->itemState), HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));

		return TRUE;
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnInputLangChange(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	IMEModeConstants ime = GetEffectiveIMEMode();
	if((ime != imeNoControl) && (GetFocus() == *this)) {
		// we've the focus, so configure the IME
		SetCurrentIMEContextMode(*this, ime);
	} else {
		ime = GetEffectiveIMEMode_Edit();
		if((ime != imeNoControl) && (GetFocus() == containedEdit.m_hWnd)) {
			// we've the focus, so configure the IME
			SetCurrentIMEContextMode(containedEdit.m_hWnd, ime);
		}
	}
	return lr;
}

LRESULT ExplorerListView::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 0;
	// required by OnKeyDownNotification
	lastKeyDownLParam = lParam;

	if(!(properties.disabledEvents & deListKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}

	wasHandled = FALSE;
	if(cachedSettings.hStateImageList && (GetAsyncKeyState(VK_SPACE) & 0x8000)) {
		// toggle the caret's state image
		LVITEMINDEX caretItem = {-1, 0};
		if(IsComctl32Version610OrNewer()) {
			SendMessage(LVM_GETNEXTITEMINDEX, reinterpret_cast<WPARAM>(&caretItem), MAKELPARAM(LVNI_FOCUSED, 0));
		} else {
			caretItem.iItem = static_cast<int>(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_FOCUSED, 0)));
		}
		if(caretItem.iItem != -1) {
			// retrieve the old state image's index
			ULONG state = 0;
			HRESULT hr = E_FAIL;
			CComPtr<IListView_WIN7> pListView7 = NULL;
			CComPtr<IListView_WINVISTA> pListViewVista = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
				ATLASSUME(pListView7);
				hr = pListView7->GetItemState(caretItem.iItem, 0, LVIS_STATEIMAGEMASK, &state);
			} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
				ATLASSUME(pListViewVista);
				hr = pListViewVista->GetItemState(caretItem.iItem, 0, LVIS_STATEIMAGEMASK, &state);
			} else {
				state = static_cast<ULONG>(SendMessage(LVM_GETITEMSTATE, caretItem.iItem, LVIS_STATEIMAGEMASK));
				hr = S_OK;
			}
			int previousStateImage = (state & LVIS_STATEIMAGEMASK) >> 12;
			// How many state images do we have?
			int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
			/* NOTE: While the treeview's state imagelist contains 3 icons, the listview's contains only 2.
			         However, the indexes are still one-based.
			   TODO: Check whether this is true for all versions of Windows. */
			// calc the new state image
			int newStateImage = (previousStateImage + 1) % (c + 1);
			if(newStateImage == 0) {
				// state images are one-based
				newStateImage = 1;
			}
			int i = newStateImage;

			CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(caretItem, this);
			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_ItemStateImageChanging(pLvwItem, previousStateImage, reinterpret_cast<PLONG>(&newStateImage), siccbKeyboard, &cancel);
			if(cancel != VARIANT_FALSE) {
				wasHandled = TRUE;
				return 0;
			}

			wasHandled = TRUE;
			lr = DefWindowProc(message, wParam, lParam);
			if(i != newStateImage) {
				// explicitly set the new state image
				LVITEM item = {0};
				item.state = INDEXTOSTATEIMAGEMASK(newStateImage);
				item.stateMask = LVIS_STATEIMAGEMASK;
				if(IsComctl32Version610OrNewer() && SendMessage(LVM_ISGROUPVIEWENABLED, 0, 0) && (GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
					if(SendMessage(LVM_SETITEMINDEXSTATE, reinterpret_cast<WPARAM>(&caretItem), reinterpret_cast<LPARAM>(&item))) {
						Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImage, siccbKeyboard);
					}
				} else {
					if(SendMessage(LVM_SETITEMSTATE, caretItem.iItem, reinterpret_cast<LPARAM>(&item))) {
						Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImage, siccbKeyboard);
					}
				}
			} else {
				// just raise the event
				Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImage, siccbKeyboard);
			}
		}
	}

	if(!wasHandled) {
		wasHandled = TRUE;
		lr = DefWindowProc(message, wParam, lParam);
	}

	return lr;
}

LRESULT ExplorerListView::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnLButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(cachedSettings.hStateImageList) {
		UINT hitTestDetails = 0;
		// retrieve the item whose bounding rectangle the double-click was within
		LVITEMINDEX itemStateImageClicked;
		HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestDetails, &itemStateImageClicked, NULL, TRUE, FALSE, TRUE);
		// NOTE: The following line is NO typo.
		if(hitTestDetails == LVHT_ONITEMSTATEICON) {
			// the item's state image was clicked

			VARIANT_BOOL cancel = VARIANT_FALSE;
			// the item's state image will be toggled - retrieve the previous state image's index
			UINT state = static_cast<UINT>(SendMessage(LVM_GETITEMSTATE, itemStateImageClicked.iItem, LVIS_STATEIMAGEMASK));
			int previousStateImage = (state & LVIS_STATEIMAGEMASK) >> 12;
			// How many state images do we have?
			int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
			/* NOTE: While the treeview's state image list contains 3 icons, the listview's contains only 2.
			         However, the indexes are still one-based.
			   TODO: Check whether this is true for all versions of Windows. */
			// calc the new state image
			int newStateImage = (previousStateImage + 1) % (c + 1);
			if(newStateImage == 0) {
				// state images are one-based
				newStateImage = 1;
			}
			int newStateImageToUse = newStateImage;

			CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemStateImageClicked, this);
			Raise_ItemStateImageChanging(pLvwItem, previousStateImage, reinterpret_cast<PLONG>(&newStateImageToUse), siccbMouse, &cancel);
			if(cancel != VARIANT_FALSE) {
				newStateImageToUse = previousStateImage;
			}

			LRESULT lr = DefWindowProc(message, wParam, lParam);
			wasHandled = TRUE;

			// finish the state image change
			if(newStateImageToUse != newStateImage) {
				LVITEM item = {0};
				item.state = INDEXTOSTATEIMAGEMASK(newStateImageToUse);
				item.stateMask = LVIS_STATEIMAGEMASK;
				if(SendMessage(LVM_SETITEMSTATE, itemStateImageClicked.iItem, reinterpret_cast<LPARAM>(&item))) {
					if(cancel == VARIANT_FALSE) {
						Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImageToUse, siccbMouse);
					}
				}
			} else {
				// just raise the event
				if(cancel == VARIANT_FALSE) {
					Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImageToUse, siccbMouse);
				}
			}

			return lr;
		}
	}

	return 0;
}

LRESULT ExplorerListView::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	UINT hitTestDetails = 0;
	// retrieve the item whose bounding rectangle the click was within
	LVITEMINDEX itemStateImageClicked;
	HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestDetails, &itemStateImageClicked, NULL, TRUE, FALSE, TRUE);
	// NOTE: The following line is NO typo.
	if(hitTestDetails != LVHT_ONITEMSTATEICON) {
		// the item's state image was not clicked
		itemStateImageClicked.iItem = -1;
	}

	int previousStateImage = 0;
	int newStateImage = 0;
	int newStateImageToUse = newStateImage;
	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(itemStateImageClicked.iItem != -1 && cachedSettings.hStateImageList) {
		// the item's state image will be toggled - retrieve the previous state image's index
		UINT state = static_cast<UINT>(SendMessage(LVM_GETITEMSTATE, itemStateImageClicked.iItem, LVIS_STATEIMAGEMASK));
		previousStateImage = (state & LVIS_STATEIMAGEMASK) >> 12;
		// How many state images do we have?
		int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
		/* NOTE: While the treeview's state image list contains 3 icons, the listview's contains only 2.
		         However, the indexes are still one-based.
		   TODO: Check whether this is true for all versions of Windows. */
		// calc the new state image
		newStateImage = (previousStateImage + 1) % (c + 1);
		if(newStateImage == 0) {
			// state images are one-based
			newStateImage = 1;
		}
		newStateImageToUse = newStateImage;

		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemStateImageClicked, this);
		Raise_ItemStateImageChanging(pLvwItem, previousStateImage, reinterpret_cast<PLONG>(&newStateImageToUse), siccbMouse, &cancel);
		if(cancel != VARIANT_FALSE) {
			newStateImageToUse = previousStateImage;
		}
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(itemStateImageClicked.iItem != -1 && cachedSettings.hStateImageList) {
		// finish the state image change
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemStateImageClicked, this);
		if(newStateImageToUse != newStateImage) {
			LVITEM item = {0};
			item.state = INDEXTOSTATEIMAGEMASK(newStateImageToUse);
			item.stateMask = LVIS_STATEIMAGEMASK;
			if(SendMessage(LVM_SETITEMSTATE, itemStateImageClicked.iItem, reinterpret_cast<LPARAM>(&item))) {
				if(cancel == VARIANT_FALSE) {
					Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImageToUse, siccbMouse);
				}
			}
		} else {
			// just raise the event
			if(cancel == VARIANT_FALSE) {
				Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImageToUse, siccbMouse);
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	mouseStatus.lastMouseUpButton = 1/*MouseButtonConstants.vbLeftButton*/;
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerListView::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	mouseStatus.lastMouseUpButton = 4/*MouseButtonConstants.vbMiddleButton*/;
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnMouseActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* Hide the message from CComControl, so that the control doesn't go into label-editing mode if the
	   caret item was clicked. */
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = !cachedSettings.hotTracking;

	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	return 0;
}

LRESULT ExplorerListView::OnMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	if(dragDropStatus.IsDragging()) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_DragMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	LVITEMINDEX previousHotItem = cachedSettings.hotItem;
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(!(properties.disabledEvents & deHotItemChangeEvents)) {
		if(cachedSettings.hotItem.iItem != previousHotItem.iItem || cachedSettings.hotItem.iGroup != previousHotItem.iGroup) {
			CComPtr<IListViewItem> pPreviousHotItem = ClassFactory::InitListItem(previousHotItem, this);
			CComPtr<IListViewItem> pNewHotItem = ClassFactory::InitListItem(cachedSettings.hotItem, this);
			Raise_HotItemChanged(pPreviousHotItem, pNewHotItem);
		} else if(cachedSettings.hotItemTmp.iItem != previousHotItem.iItem || cachedSettings.hotItemTmp.iGroup != previousHotItem.iGroup) {
			cachedSettings.hotItem = cachedSettings.hotItemTmp;
			CComPtr<IListViewItem> pPreviousHotItem = ClassFactory::InitListItem(previousHotItem, this);
			CComPtr<IListViewItem> pNewHotItem = ClassFactory::InitListItem(cachedSettings.hotItem, this);
			Raise_HotItemChanged(pPreviousHotItem, pNewHotItem);
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		ScreenToClient(&mousePosition);
		Raise_MouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerListView::OnNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			if(reinterpret_cast<LPNMHDR>(lParam)->hwndFrom == containedSysHeader32.m_hWnd) {
				return OnHeaderCustomDrawNotification(message, wParam, lParam, wasHandled);
			}
			break;

		case TTN_GETDISPINFOA:
			return OnToolTipGetDispInfoNotificationA(message, wParam, lParam, wasHandled);
			break;

		case TTN_GETDISPINFOW:
			return OnToolTipGetDispInfoNotificationW(message, wParam, lParam, wasHandled);
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ExplorerListView::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnParentNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 0;
	wasHandled = FALSE;

	switch(wParam) {
		case WM_CREATE: {
			lr = DefWindowProc(message, wParam, lParam);
			wasHandled = TRUE;

			TCHAR pBuffer[MAX_PATH + 1];
			GetClassName(reinterpret_cast<HWND>(lParam), pBuffer, MAX_PATH);
			if(lstrcmp(pBuffer, reinterpret_cast<LPCTSTR>(&WC_EDIT)) == 0) {
				Raise_CreatedEditControlWindow(reinterpret_cast<HWND>(lParam));
			} else if(lstrcmp(pBuffer, reinterpret_cast<LPCTSTR>(&WC_HEADER)) == 0) {
				containedSysHeader32.SubclassWindow(reinterpret_cast<HWND>(lParam));
				SetTimer(timers.ID_CREATED, timers.INT_CREATED);
				//Raise_CreatedHeaderControlWindow(reinterpret_cast<HWND>(lParam));
			}
			break;
		}
	}
	return lr;
}

LRESULT ExplorerListView::OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	UINT hitTestDetails = 0;
	// retrieve the item whose bounding rectangle the click was within
	LVITEMINDEX itemStateImageClicked;
	HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestDetails, &itemStateImageClicked, NULL, TRUE, FALSE, TRUE);
	// NOTE: The following line is NO typo.
	if(hitTestDetails != LVHT_ONITEMSTATEICON) {
		// the item's state image was not clicked
		itemStateImageClicked.iItem = -1;
	}

	int previousStateImage = 0;
	int newStateImage = 0;
	int newStateImageToUse = newStateImage;
	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(itemStateImageClicked.iItem != -1 && cachedSettings.hStateImageList) {
		// the item's state image will be toggled - retrieve the previous state image's index
		UINT state = static_cast<UINT>(SendMessage(LVM_GETITEMSTATE, itemStateImageClicked.iItem, LVIS_STATEIMAGEMASK));
		previousStateImage = (state & LVIS_STATEIMAGEMASK) >> 12;
		// How many state images do we have?
		int c = ImageList_GetImageCount(cachedSettings.hStateImageList);
		/* NOTE: While the treeview's state image list contains 3 icons, the listview's contains only 2.
		         However, the indexes are still one-based.
		   TODO: Check whether this is true for all versions of Windows. */
		// calc the new state image
		newStateImage = (previousStateImage + 1) % (c + 1);
		if(newStateImage == 0) {
			// state images are one-based
			newStateImage = 1;
		}
		newStateImageToUse = newStateImage;

		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemStateImageClicked, this);
		Raise_ItemStateImageChanging(pLvwItem, previousStateImage, reinterpret_cast<PLONG>(&newStateImageToUse), siccbMouse, &cancel);
		if(cancel != VARIANT_FALSE) {
			newStateImageToUse = previousStateImage;
		}
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(itemStateImageClicked.iItem != -1 && cachedSettings.hStateImageList) {
		// finish the state image change
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemStateImageClicked, this);
		if(newStateImageToUse != newStateImage) {
			LVITEM item = {0};
			item.state = INDEXTOSTATEIMAGEMASK(newStateImageToUse);
			item.stateMask = LVIS_STATEIMAGEMASK;
			if(SendMessage(LVM_SETITEMSTATE, itemStateImageClicked.iItem, reinterpret_cast<LPARAM>(&item))) {
				if(cancel == VARIANT_FALSE) {
					Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImageToUse, siccbMouse);
				}
			}
		} else {
			// just raise the event
			if(cancel == VARIANT_FALSE) {
				Raise_ItemStateImageChanged(pLvwItem, previousStateImage, newStateImageToUse, siccbMouse);
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	flags.receivedRButtonUp = TRUE;

	mouseStatus.lastMouseUpButton = 2/*MouseButtonConstants.vbRightButton*/;
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the listview?
	RECT clientArea = {0};
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(PtInRect(&clientArea, mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(IsComctl32Version600()) {
		/* HACK: Custom draw won't work in the first repaint after the control has received the focus. So we
		         handle the focus change ourselves and redraw immediately. */
		if(m_bInPlaceActive) {
			CComPtr<IOleObject> pOleObject;
			ControlQueryInterface(__uuidof(IOleObject), reinterpret_cast<LPVOID*>(&pOleObject));
			if(pOleObject) {
				pOleObject->DoVerb(OLEIVERB_UIACTIVATE, NULL, m_spClientSite, 0, m_hWndCD, &m_rcPos);
			}
			CComQIPtr<IOleControlSite, &__uuidof(IOleControlSite)> pSite(m_spClientSite);
			if(m_bInPlaceActive && pSite) {
				pSite->OnFocus(TRUE);
			}
		}

		Invalidate();
		return DefWindowProc(message, wParam, lParam);
	} else if(IsComctl32Version610OrNewer() && containedSysHeader32.IsWindowVisible()) {
		// insert the header control into the tab order
		if(reinterpret_cast<HWND>(wParam) != containedSysHeader32.m_hWnd) {
			if((GetKeyState(VK_TAB) & 0x8000) && (GetKeyState(VK_SHIFT) & 0x8000)) {
				// both pressed
				if(m_bInPlaceActive) {
					CComPtr<IOleObject> pOleObject;
					ControlQueryInterface(__uuidof(IOleObject), reinterpret_cast<LPVOID*>(&pOleObject));
					if(pOleObject) {
						pOleObject->DoVerb(OLEIVERB_UIACTIVATE, NULL, m_spClientSite, 0, m_hWndCD, &m_rcPos);
					}
					CComQIPtr<IOleControlSite, &__uuidof(IOleControlSite)> pSite(m_spClientSite);
					if(m_bInPlaceActive && pSite) {
						pSite->OnFocus(TRUE);
					}
				}

				containedSysHeader32.SetFocus();
				return 0;
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_EXLVW_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to create a new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_FONT);
	}

	return lr;
}

LRESULT ExplorerListView::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_CREATED:
			KillTimer(timers.ID_CREATED);
			Raise_CreatedHeaderControlWindow(containedSysHeader32);
			break;

		case timers.ID_REDRAW:
			KillTimer(timers.ID_REDRAW);
			hBuiltInStateImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_STATE, 0));
			cachedSettings.hStateImageList = hBuiltInStateImageList;
			if(IsComctl32Version610OrNewer()) {
				hBuiltInHeaderStateImageList = reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_STATE, 0));
				cachedSettings.hHeaderStateImageList = hBuiltInHeaderStateImageList;
			}
			RedrawWindow();
			break;

		case timers.ID_DRAGSCROLL:
			AutoScroll();
			break;

		case timers.ID_CARETCHANGED: {
			KillTimer(timers.ID_CARETCHANGED);
			LVITEMINDEX itemIndex = {-1, 0};
			if(IsComctl32Version610OrNewer() && SendMessage(LVM_GETNEXTITEMINDEX, reinterpret_cast<WPARAM>(&itemIndex), MAKELPARAM(LVNI_FOCUSED, 0)) && itemIndex.iItem == -1) {
				// the caret was moved from 'caretChangedStatus.previousCaretItem' to -1
				caretChangedStatus.previousCaretItem = caretChangedStatus.newCaretItem;
				caretChangedStatus.newCaretItem.iItem = -1;
				caretChangedStatus.newCaretItem.iGroup = 0;
			} else if(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_FOCUSED, 0)) == -1) {
				// the caret was moved from 'caretChangedStatus.previousCaretItem' to -1
				caretChangedStatus.previousCaretItem = caretChangedStatus.newCaretItem;
				caretChangedStatus.newCaretItem.iItem = -1;
				caretChangedStatus.newCaretItem.iGroup = 0;
			}
			Raise_CaretChanged(caretChangedStatus.previousCaretItem, caretChangedStatus.newCaretItem);
			break;
		}

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ExplorerListView::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ExplorerListView::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			mouseStatus.lastMouseUpButton = embXButton1;
			break;
		case XBUTTON2:
			mouseStatus.lastMouseUpButton = embXButton2;
			break;
	}
	if(!(properties.disabledEvents & deListMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		button = mouseStatus.lastMouseUpButton;
		Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	ATLASSERT(pDetails->CtlType == ODT_LISTVIEW);
	ATLASSERT(pDetails->itemAction == ODA_DRAWENTIRE);
	ATLASSERT((pDetails->itemState & (ODS_GRAYED | ODS_DISABLED | ODS_CHECKED | ODS_DEFAULT | ODS_COMBOBOXEDIT | ODS_INACTIVE)) == 0);

	if(pDetails->hwndItem == *this) {
		LVITEMINDEX itemIndex = {pDetails->itemID, 0};
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
		Raise_OwnerDrawItem(pLvwItem, static_cast<OwnerDrawItemStateConstants>(pDetails->itemState), HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnReflectedMeasureItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPMEASUREITEMSTRUCT pDetails = reinterpret_cast<LPMEASUREITEMSTRUCT>(lParam);
	if(pDetails->CtlType == ODT_LISTVIEW) {
		pDetails->itemHeight = properties.itemHeight;
		return TRUE;
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnReflectedNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			if(reinterpret_cast<LPNMHDR>(lParam)->hwndFrom == *this) {
				return OnCustomDrawNotification(message, wParam, lParam, wasHandled);
			}
			break;
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ExplorerListView::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == NF_QUERY) {
		#ifdef UNICODE
			return NFR_UNICODE;
		#else
			return NFR_ANSI;
		#endif
	}
	return 0;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	LRESULT ExplorerListView::OnHandshake(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
	{
		LPSHBHANDSHAKE pDetails = reinterpret_cast<LPSHBHANDSHAKE>(lParam);
		ATLASSERT_POINTER(pDetails, SHBHANDSHAKE);
		if(!pDetails) {
			return FALSE;
		}

		pDetails->processed = TRUE;
		if(pDetails->structSize < CCSIZEOF_STRUCT(SHBHANDSHAKE, pCtlInterface)) {
			return FALSE;
		}

		if(pDetails->pMessageHook && shellBrowserInterface.pMessageListener && shellBrowserInterface.pInternalMessageListener) {
			// there's already a ShellBrowser attached - fail
			pDetails->errorCode = E_FAIL;
			return FALSE;
		}

		if(pDetails->pCtlBuildNumber) {
			*(pDetails->pCtlBuildNumber) = VERSION_BUILD;
		}
		if(pDetails->pCtlSupportsUnicode) {
			#ifdef UNICODE
				*(pDetails->pCtlSupportsUnicode) = TRUE;
			#else
				*(pDetails->pCtlSupportsUnicode) = FALSE;
			#endif
		}
		if(pDetails->pCtlInterface) {
			pDetails->errorCode = QueryInterface(IID_IUnknown, pDetails->pCtlInterface);
		} else {
			pDetails->errorCode = S_OK;
		}
		if(SUCCEEDED(pDetails->errorCode)) {
			shellBrowserInterface.shellBrowserBuildNumber = pDetails->shellBrowserBuildNumber;
			shellBrowserInterface.pMessageListener = pDetails->pMessageHook;
			shellBrowserInterface.pInternalMessageListener = pDetails->pShBInternalMessageHook;
			if(pDetails->ppCtlInternalMessageHook) {
				*(pDetails->ppCtlInternalMessageHook) = this;
			}
		}
		return SUCCEEDED(pDetails->errorCode);
	}
#endif

LRESULT ExplorerListView::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.pDraggedItems) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = CreateVistaOLEDragImage(dragDropStatus.pDraggedItems, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.pDraggedItems, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT ExplorerListView::OnCreateDragImage(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LVITEMINDEX itemIndex = {static_cast<int>(lParam), 0};
	return reinterpret_cast<LRESULT>(CreateLegacyDragImage(itemIndex, reinterpret_cast<LPPOINT>(wParam), NULL));
}

LRESULT ExplorerListView::OnDeleteColumn(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	LONG columnIDBeingRemoved = ColumnIndexToID(static_cast<int>(wParam));

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IListViewColumn> pLvwColumnItf = NULL;
	CComObject<ListViewColumn>* pLvwColumnObj = NULL;
	if(!(properties.disabledEvents & deColumnDeletionEvents)) {
		CComObject<ListViewColumn>::CreateInstance(&pLvwColumnObj);
		pLvwColumnObj->AddRef();
		pLvwColumnObj->SetOwner(this);
		pLvwColumnObj->Attach(static_cast<int>(wParam));
		pLvwColumnObj->QueryInterface(IID_IListViewColumn, reinterpret_cast<LPVOID*>(&pLvwColumnItf));
		pLvwColumnObj->Release();
		Raise_RemovingColumn(pLvwColumnItf, &cancel);
	}

	if(cancel == VARIANT_FALSE) {
		CComPtr<IVirtualListViewColumn> pVLvwColumnItf = NULL;
		if(!(properties.disabledEvents & deColumnDeletionEvents)) {
			CComObject<VirtualListViewColumn>* pVLvwColumnObj = NULL;
			CComObject<VirtualListViewColumn>::CreateInstance(&pVLvwColumnObj);
			pVLvwColumnObj->AddRef();
			pVLvwColumnObj->SetOwner(this);
			if(pLvwColumnObj) {
				pLvwColumnObj->SaveState(pVLvwColumnObj);
			}

			pVLvwColumnObj->QueryInterface(IID_IVirtualListViewColumn, reinterpret_cast<LPVOID*>(&pVLvwColumnItf));
			pVLvwColumnObj->Release();
		}

		if(!(properties.disabledEvents & deHeaderMouseEvents)) {
			if(static_cast<int>(wParam) == columnUnderMouse) {
				// we're removing the column below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				UINT hitTestDetails = 0;
				HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				Raise_ColumnMouseLeave(pLvwColumnItf, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
				columnUnderMouse = -1;
			}
		}

		if(!(properties.disabledEvents & deFreeColumnData)) {
			if(pLvwColumnItf) {
				CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(static_cast<int>(wParam), this, FALSE);
				Raise_FreeColumnData(pLvwColumn);
			} else {
				Raise_FreeColumnData(pLvwColumnItf);
			}
		}

		// finally pass the message to the header
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			if(static_cast<int>(wParam) == dragDropStatus.draggedColumn) {
				dragDropStatus.draggedColumn = -1;
			}
			if(!(properties.disabledEvents & deColumnDeletionEvents)) {
				Raise_RemovedColumn(pVLvwColumnItf);
			}
		}

		#ifdef USE_STL
			columnIndexes.erase(columnIDBeingRemoved);
			for(std::unordered_map< LONG, int, ItemIndexHasher<LONG> >::iterator iter = columnIndexes.begin(); iter != columnIndexes.end(); ++iter) {
				if(iter->second > static_cast<int>(wParam)) {
					--(iter->second);
				}
			}
			std::list<ColumnData>::iterator iter2 = columnParams.begin();
			if(iter2 != columnParams.end()) {
				std::advance(iter2, static_cast<int>(wParam));
				if(iter2 != columnParams.end()) {
					columnParams.erase(iter2);
				}
			}
		#else
			columnIndexes.RemoveKey(columnIDBeingRemoved);
			POSITION p = columnIndexes.GetStartPosition();
			while(p) {
				CAtlMap<LONG, int>::CPair* pPair = columnIndexes.GetAt(p);
				if(pPair->m_value > static_cast<int>(wParam)) {
					--(pPair->m_value);
				}
				columnIndexes.GetNext(p);
			}
			p = columnParams.FindIndex(static_cast<int>(wParam));
			if(p) {
				columnParams.RemoveAt(p);
			}
		#endif
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pInternalMessageListener) {
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_FREECOLUMN, columnIDBeingRemoved, 0);
			}
		#endif

		if(!(properties.disabledEvents & deHeaderMouseEvents)) {
			if(lr) {
				// maybe we have a new column below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				UINT hitTestDetails = 0;
				int newColumnUnderMouse = HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				if(newColumnUnderMouse != columnUnderMouse) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(columnUnderMouse != -1) {
						pLvwColumnItf = ClassFactory::InitColumn(columnUnderMouse, this);
						Raise_ColumnMouseLeave(pLvwColumnItf, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
					}
					columnUnderMouse = newColumnUnderMouse;
					if(columnUnderMouse != -1) {
						pLvwColumnItf = ClassFactory::InitColumn(columnUnderMouse, this);
						Raise_ColumnMouseEnter(pLvwColumnItf, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
					}
				}
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnDeleteAllItems(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		Raise_RemovingItem(NULL, &cancel);
	}
	if(cancel == VARIANT_FALSE) {
		if(!(properties.disabledEvents & deHotItemChangeEvents)) {
			if(cachedSettings.hotItem.iItem != -1) {
				// we're removing the hot item
				CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(cachedSettings.hotItem, this);
				Raise_HotItemChanged(pLvwItem, NULL);
				cachedSettings.hotItem.iItem = -1;
				cachedSettings.hotItem.iGroup = 0;
			}
		}

		if(!(properties.disabledEvents & deListMouseEvents)) {
			if(itemUnderMouse.iItem != -1) {
				// we're removing the item below the mouse cursor
				CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);

				UINT h = 0;
				int dummy = -1;
				// we use dummy to make sure LVM_SUBITEMHITTEST is used
				HitTest(mousePosition.x, mousePosition.y, &h, NULL, &dummy, TRUE, TRUE, FALSE);
				HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
				if(subItemUnderMouse != -1) {
					// we're removing the sub-item below the mouse cursor
					CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
					Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					subItemUnderMouse = -1;
				}
				Raise_ItemMouseLeave(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				itemUnderMouse.iItem = -1;
				itemUnderMouse.iGroup = 0;
			}

			LVITEMINDEX caretItem = {-1, 0};
			if(!IsComctl32Version610OrNewer() || !SendMessage(LVM_GETNEXTITEMINDEX, reinterpret_cast<WPARAM>(&caretItem), MAKELPARAM(LVNI_FOCUSED, 0))) {
				caretItem.iItem = static_cast<int>(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_FOCUSED, 0)));
			}
			if(caretItem.iItem != -1) {
				LVITEMINDEX dummy = {-1, 0};
				Raise_CaretChanged(caretItem, dummy);
			}
		}

		// finally pass the message to the listview
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			if(!(properties.disabledEvents & deItemDeletionEvents)) {
				Raise_RemovedItem(NULL);
			}
		}

		RemoveItemFromItemContainers(-1);
		#ifdef USE_STL
			itemIDs.clear();
			itemParams.clear();
			groups.clear();
		#else
			itemIDs.RemoveAll();
			itemParams.RemoveAll();
			groups.RemoveAll();
		#endif
	}

	return lr;
}

LRESULT ExplorerListView::OnDeleteItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IListViewItem> pLvwItemItf = NULL;
	CComObject<ListViewItem>* pLvwItemObj = NULL;
	if(!(properties.disabledEvents & deItemDeletionEvents)) {
		CComObject<ListViewItem>::CreateInstance(&pLvwItemObj);
		pLvwItemObj->AddRef();
		pLvwItemObj->SetOwner(this);
		LVITEMINDEX itemIndex = {static_cast<int>(wParam), 0};
		pLvwItemObj->Attach(itemIndex);
		pLvwItemObj->QueryInterface(IID_IListViewItem, reinterpret_cast<LPVOID*>(&pLvwItemItf));
		pLvwItemObj->Release();
		Raise_RemovingItem(pLvwItemItf, &cancel);
	}

	if(cancel == VARIANT_FALSE) {
		CComPtr<IVirtualListViewItem> pVLvwItemItf = NULL;
		if(!(properties.disabledEvents & deItemDeletionEvents)) {
			CComObject<VirtualListViewItem>* pVLvwItemObj = NULL;
			CComObject<VirtualListViewItem>::CreateInstance(&pVLvwItemObj);
			pVLvwItemObj->AddRef();
			pVLvwItemObj->SetOwner(this);
			if(pLvwItemObj) {
				pLvwItemObj->SaveState(pVLvwItemObj);
			}

			pVLvwItemObj->QueryInterface(IID_IVirtualListViewItem, reinterpret_cast<LPVOID*>(&pVLvwItemItf));
			pVLvwItemObj->Release();
		}

		if(!(properties.disabledEvents & deHotItemChangeEvents)) {
			if(static_cast<int>(wParam) == cachedSettings.hotItem.iItem) {
				// we're removing the hot item
				CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(cachedSettings.hotItem, this);
				Raise_HotItemChanged(pLvwItem, NULL);
				cachedSettings.hotItem.iItem = -1;
				cachedSettings.hotItem.iGroup = 0;
			}
		}

		if(!(properties.disabledEvents & deListMouseEvents)) {
			if(static_cast<int>(wParam) == itemUnderMouse.iItem) {
				// we're removing the item below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);

				UINT h = 0;
				int dummy = -1;
				// we use dummy to make sure LVM_SUBITEMHITTEST is used
				HitTest(mousePosition.x, mousePosition.y, &h, NULL, &dummy, TRUE, TRUE, FALSE);
				HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
				if(subItemUnderMouse != -1) {
					// we're removing the sub-item below the mouse cursor
					CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
					Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					subItemUnderMouse = -1;
				}
				Raise_ItemMouseLeave(pLvwItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				itemUnderMouse.iItem = -1;
				itemUnderMouse.iGroup = 0;
			}
		}

		// finally pass the message to the listview
		flags.itemIDBeingRemoved = ItemIndexToID(static_cast<int>(wParam));
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			if(!(properties.disabledEvents & deItemDeletionEvents)) {
				Raise_RemovedItem(pVLvwItemItf);
			}
		}

		if(!(properties.disabledEvents & deListMouseEvents)) {
			if(lr) {
				// maybe we have a new (sub-)item below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);

				UINT h = 0;
				LVITEMINDEX newItemUnderMouse;
				int newSubItemUnderMouse = -1;
				HitTest(mousePosition.x, mousePosition.y, &h, &newItemUnderMouse, &newSubItemUnderMouse);
				HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
				if(newItemUnderMouse.iItem != itemUnderMouse.iItem || newItemUnderMouse.iGroup != itemUnderMouse.iGroup) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(itemUnderMouse.iItem != -1) {
						pLvwItemItf = ClassFactory::InitListItem(itemUnderMouse, this);
						if(subItemUnderMouse > 0) {
							CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
							Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						}
						Raise_ItemMouseLeave(pLvwItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}

					itemUnderMouse = newItemUnderMouse;
					subItemUnderMouse = newSubItemUnderMouse;

					if(itemUnderMouse.iItem != -1) {
						hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
						pLvwItemItf = ClassFactory::InitListItem(itemUnderMouse, this);
						Raise_ItemMouseEnter(pLvwItemItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						if(subItemUnderMouse > 0) {
							CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
							Raise_SubItemMouseEnter(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						}
					}
				}
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnFindItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LVFINDINFO* pFindInfo = reinterpret_cast<LVFINDINFO*>(lParam);
	#ifdef DEBUG
		if(message == LVM_FINDITEMA) {
			ATLASSERT_POINTER(pFindInfo, LVFINDINFOA);
		} else {
			ATLASSERT_POINTER(pFindInfo, LVFINDINFOW);
		}
	#endif
	if(!pFindInfo) {
		wasHandled = FALSE;
		return -1;
	}

	if(((pFindInfo->flags & LVFI_PARAM) == LVFI_PARAM) && !RunTimeHelper::IsCommCtrl6()) {
		// SysListView32 can't handle the message
		int i = static_cast<int>(wParam);
		#ifdef USE_STL
			for(size_t j = itemParams.size(); j > 0; --j) {
				++i;
				if(i == static_cast<int>(itemParams.size())) {
					if(pFindInfo->flags & LVFI_WRAP) {
						i = 0;
					} else {
						break;
					}
				}
				std::list<ItemData>::iterator iter = itemParams.begin();
				std::advance(iter, i);
				if(iter != itemParams.end()) {
					if(iter->lParam == pFindInfo->lParam) {
						return i;
					}
				}
			}
		#else
			for(size_t j = itemParams.GetCount(); j > 0; --j) {
				++i;
				if(i == static_cast<int>(itemParams.GetCount())) {
					if(pFindInfo->flags & LVFI_WRAP) {
						i = 0;
					} else {
						break;
					}
				}
				POSITION p = itemParams.FindIndex(i);
				if(p) {
					if(itemParams.GetAt(p).lParam == pFindInfo->lParam) {
						return i;
					}
				}
			}
		#endif
		return -1;
	}

	// let SysListView32 handle the message
	wasHandled = FALSE;
	return -1;
}

LRESULT ExplorerListView::OnGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL overwriteLParam = FALSE;

	LPLVITEM pItemData = reinterpret_cast<LPLVITEM>(lParam);
	#ifdef DEBUG
		if(message == LVM_GETITEMA) {
			ATLASSERT_POINTER(pItemData, LVITEMA);
		} else {
			ATLASSERT_POINTER(pItemData, LVITEMW);
		}
	#endif
	if(!pItemData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pItemData->mask & LVIF_NOINTERCEPTION) {
		// SysListView32 shouldn't see this flag
		pItemData->mask &= ~LVIF_NOINTERCEPTION;
	} else if(pItemData->mask == LVIF_PARAM) {
		if(!RunTimeHelper::IsCommCtrl6()) {
			// just look up in the 'itemParams' list
			#ifdef USE_STL
				std::list<ItemData>::iterator iter = itemParams.begin();
				if(iter != itemParams.end()) {
					std::advance(iter, pItemData->iItem);
					if(iter != itemParams.end()) {
						pItemData->lParam = iter->lParam;
						return TRUE;
					}
				}
			#else
				POSITION p = itemParams.FindIndex(pItemData->iItem);
				if(p) {
					pItemData->lParam = itemParams.GetAt(p).lParam;
					return TRUE;
				}
			#endif
			// item not found
			return FALSE;
		}
	} else if(pItemData->mask & LVIF_PARAM) {
		if(!RunTimeHelper::IsCommCtrl6()) {
			overwriteLParam = TRUE;
		}
	}

	// forward the message
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(overwriteLParam) {
		#ifdef USE_STL
			std::list<ItemData>::iterator iter = itemParams.begin();
			if(iter != itemParams.end()) {
				std::advance(iter, pItemData->iItem);
				if(iter != itemParams.end()) {
					pItemData->lParam = iter->lParam;
				}
			}
		#else
			POSITION p = itemParams.FindIndex(pItemData->iItem);
			if(p) {
				pItemData->lParam = itemParams.GetAt(p).lParam;
			}
		#endif
	}

	return lr;
}

LRESULT ExplorerListView::OnInsertColumn(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	ATLASSERT_POINTER(reinterpret_cast<LPLVCOLUMN>(lParam), LVCOLUMN);

	if(lParam) {
		bufferedColumnData = *reinterpret_cast<LPLVCOLUMN>(lParam);
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnInsertGroup(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	int insertedGroup = -1;

	if(!(properties.disabledEvents & deGroupInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;

		CComObject<VirtualListViewGroup>* pVLvwGroupObj = NULL;
		CComPtr<IVirtualListViewGroup> pVLvwGroupItf = NULL;
		CComObject<VirtualListViewGroup>::CreateInstance(&pVLvwGroupObj);
		pVLvwGroupObj->AddRef();
		pVLvwGroupObj->SetOwner(this);
		pVLvwGroupObj->LoadState(reinterpret_cast<PLVGROUP>(lParam));
		pVLvwGroupObj->QueryInterface(IID_IVirtualListViewGroup, reinterpret_cast<LPVOID*>(&pVLvwGroupItf));
		pVLvwGroupObj->Release();

		Raise_InsertingGroup(pVLvwGroupItf, &cancel);
		pVLvwGroupObj = NULL;

		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the listview
			insertedGroup = static_cast<int>(DefWindowProc(message, wParam, lParam));
			if(insertedGroup != -1) {
				#ifdef USE_STL
					std::vector<int>::iterator iter = groups.begin();
					if(iter != groups.end()) {
						std::advance(iter, insertedGroup);
					}
					groups.insert(iter, reinterpret_cast<PLVGROUP>(lParam)->iGroupId);
				#else
					groups.InsertAt(insertedGroup, reinterpret_cast<PLVGROUP>(lParam)->iGroupId);
				#endif

				CComPtr<IListViewGroup> pLvwGroupItf = ClassFactory::InitGroup(reinterpret_cast<PLVGROUP>(lParam)->iGroupId, this);
				if(pLvwGroupItf) {
					Raise_InsertedGroup(pLvwGroupItf);
				}
			}
		}
	} else {
		// finally pass the message to the listview
		insertedGroup = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedGroup != -1) {
			#ifdef USE_STL
				std::vector<int>::iterator iter = groups.begin();
				if(iter != groups.end()) {
					std::advance(iter, insertedGroup);
				}
				groups.insert(iter, reinterpret_cast<PLVGROUP>(lParam)->iGroupId);
			#else
				groups.InsertAt(insertedGroup, reinterpret_cast<PLVGROUP>(lParam)->iGroupId);
			#endif
		}
	}

	if(!(properties.disabledEvents & deListMouseEvents)) {
		if(insertedGroup != -1) {
			// maybe we have a new (sub-)item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT h = 0;
			LVITEMINDEX newItemUnderMouse;
			int newSubItemUnderMouse = -1;
			HitTest(mousePosition.x, mousePosition.y, &h, &newItemUnderMouse, &newSubItemUnderMouse);
			HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
			if(newItemUnderMouse.iItem != itemUnderMouse.iItem || newItemUnderMouse.iGroup != itemUnderMouse.iGroup) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse.iItem != -1) {
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					Raise_ItemMouseLeave(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				subItemUnderMouse = newSubItemUnderMouse;

				if(itemUnderMouse.iItem != -1) {
					hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseEnter(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return insertedGroup;
}

LRESULT ExplorerListView::OnInsertGroupSorted(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	int insertedGroup = -1;

	if(!(properties.disabledEvents & deGroupInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;

		CComObject<VirtualListViewGroup>* pVLvwGroupObj = NULL;
		CComPtr<IVirtualListViewGroup> pVLvwGroupItf = NULL;
		CComObject<VirtualListViewGroup>::CreateInstance(&pVLvwGroupObj);
		pVLvwGroupObj->AddRef();
		pVLvwGroupObj->SetOwner(this);
		pVLvwGroupObj->LoadState(&reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->lvGroup);
		pVLvwGroupObj->QueryInterface(IID_IVirtualListViewGroup, reinterpret_cast<LPVOID*>(&pVLvwGroupItf));
		pVLvwGroupObj->Release();

		Raise_InsertingGroup(pVLvwGroupItf, &cancel);
		pVLvwGroupObj = NULL;

		if(cancel == VARIANT_FALSE) {
			/* How does LVM_INSERTGROUPSORTED work? Well, the comparison function is called with the new
			   group's ID and each existing group's ID until the correct slot is found. Existing groups
			   don't get sorted! So plugging in our own comparison function is useless, because it won't
			   find the new group's ID within the 'groups' vector and one of the IDs to compare is always
			   the new one. Additionally the message's return value gives us the correct position index. */
			/*SortGroupsDetails newLParam = {0};
			newLParam.pRealComparisonFunction = reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pfnGroupCompare;
			newLParam.pAppDefinedInformation = reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pvData;
			newLParam.pGroups = &groups;

			reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pfnGroupCompare = ::MyLVGroupCompare;
			reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pvData = &newLParam;*/

			// finally pass the message to the listview
			// NOTE: According to Windows SDK, the return value is not used, but actually it's the index.
			insertedGroup = static_cast<int>(DefWindowProc(message, wParam, lParam));
			if(insertedGroup != -1) {
				#ifdef USE_STL
					std::vector<int>::iterator iter = groups.begin();
					if(iter != groups.end()) {
						std::advance(iter, insertedGroup);
					}
					groups.insert(iter, reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->lvGroup.iGroupId);
				#else
					groups.InsertAt(insertedGroup, reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->lvGroup.iGroupId);
				#endif

				CComPtr<IListViewGroup> pLvwGroupItf = ClassFactory::InitGroup(reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->lvGroup.iGroupId, this);
				if(pLvwGroupItf) {
					Raise_InsertedGroup(pLvwGroupItf);
				}
			}
		}
	} else {
		/* How does LVM_INSERTGROUPSORTED work? Well, the comparison function is called with the new
		   group's ID and each existing group's ID until the correct slot is found. Existing groups
		   don't get sorted! So plugging in our own comparison function is useless, because it won't
		   find the new group's ID within the 'groups' vector and one of the IDs to compare is always
		   the new one. Additionally the message's return value gives us the correct position index. */
		/*SortGroupsDetails newLParam = {0};
		newLParam.pRealComparisonFunction = reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pfnGroupCompare;
		newLParam.pAppDefinedInformation = reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pvData;
		newLParam.pGroups = &groups;

		reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pfnGroupCompare = ::MyLVGroupCompare;
		reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->pvData = &newLParam;*/

		// finally pass the message to the listview
		// NOTE: According to Windows SDK, the return value is not used, but actually it's the index.
		insertedGroup = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedGroup != -1) {
			#ifdef USE_STL
				std::vector<int>::iterator iter = groups.begin();
				if(iter != groups.end()) {
					std::advance(iter, insertedGroup);
				}
				groups.insert(iter, reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->lvGroup.iGroupId);
			#else
				groups.InsertAt(insertedGroup, reinterpret_cast<PLVINSERTGROUPSORTED>(wParam)->lvGroup.iGroupId);
			#endif
		}
	}

	if(!(properties.disabledEvents & deListMouseEvents)) {
		if(insertedGroup != -1) {
			// maybe we have a new (sub-)item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT h = 0;
			LVITEMINDEX newItemUnderMouse;
			int newSubItemUnderMouse = -1;
			HitTest(mousePosition.x, mousePosition.y, &h, &newItemUnderMouse, &newSubItemUnderMouse);
			HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
			if(newItemUnderMouse.iItem != itemUnderMouse.iItem || newItemUnderMouse.iGroup != itemUnderMouse.iGroup) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse.iItem != -1) {
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					Raise_ItemMouseLeave(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				subItemUnderMouse = newSubItemUnderMouse;

				if(itemUnderMouse.iItem != -1) {
					hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseEnter(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return insertedGroup;
}

LRESULT ExplorerListView::OnInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	int insertedItem = -1;

	if(!(properties.disabledEvents & deItemInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;

		CComObject<VirtualListViewItem>* pVLvwItemObj = NULL;
		CComPtr<IVirtualListViewItem> pVLvwItemItf = NULL;
		CComObject<VirtualListViewItem>::CreateInstance(&pVLvwItemObj);
		pVLvwItemObj->AddRef();
		pVLvwItemObj->SetOwner(this);

		#ifdef UNICODE
			BOOL mustConvert = (message == LVM_INSERTITEMA);
		#else
			BOOL mustConvert = (message == LVM_INSERTITEMW);
		#endif
		if(mustConvert) {
			#ifdef UNICODE
				LPLVITEMA pItemData = reinterpret_cast<LPLVITEMA>(lParam);
				LVITEMW convertedItemData = {0};
				LPSTR p = NULL;
				if(pItemData->pszText != LPSTR_TEXTCALLBACKA) {
					p = pItemData->pszText;
				}
				CA2W converter(p);
				if(pItemData->pszText == LPSTR_TEXTCALLBACKA) {
					convertedItemData.pszText = LPSTR_TEXTCALLBACKW;
				} else {
					convertedItemData.pszText = converter;
				}
			#else
				LPLVITEMW pItemData = reinterpret_cast<LPLVITEMW>(lParam);
				LVITEMA convertedItemData = {0};
				LPWSTR p = NULL;
				if(pItemData->pszText != LPSTR_TEXTCALLBACKW) {
					p = pItemData->pszText;
				}
				CW2A converter(p);
				if(pItemData->pszText == LPSTR_TEXTCALLBACKW) {
					convertedItemData.pszText = LPSTR_TEXTCALLBACKA;
				} else {
					convertedItemData.pszText = converter;
				}
			#endif
			convertedItemData.cchTextMax = pItemData->cchTextMax;
			convertedItemData.mask = pItemData->mask;
			convertedItemData.iItem = pItemData->iItem;
			convertedItemData.iSubItem = pItemData->iSubItem;
			convertedItemData.state = pItemData->state;
			convertedItemData.stateMask = pItemData->stateMask;
			convertedItemData.iImage = pItemData->iImage;
			convertedItemData.lParam = pItemData->lParam;
			convertedItemData.iIndent = pItemData->iIndent;
			// the LVITEM struct may end here, so be more careful with the remaining members
			if(convertedItemData.mask & LVIF_GROUPID) {
				convertedItemData.iGroupId = pItemData->iGroupId;
			}
			if(convertedItemData.mask & LVIF_COLUMNS) {
				convertedItemData.cColumns = pItemData->cColumns;
				convertedItemData.puColumns = pItemData->puColumns;
			}
			if(convertedItemData.mask & LVIF_COLFMT) {
				convertedItemData.cColumns = pItemData->cColumns;
				convertedItemData.piColFmt = pItemData->piColFmt;
			}
			pVLvwItemObj->LoadState(&convertedItemData);
		} else {
			pVLvwItemObj->LoadState(reinterpret_cast<LPLVITEM>(lParam));
		}

		pVLvwItemObj->QueryInterface(IID_IVirtualListViewItem, reinterpret_cast<LPVOID*>(&pVLvwItemItf));
		pVLvwItemObj->Release();

		Raise_InsertingItem(pVLvwItemItf, &cancel);
		pVLvwItemObj = NULL;

		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the listview
			LPLVITEM pDetails = reinterpret_cast<LPLVITEM>(lParam);
			pDetails->mask &= ~LVIF_NOINTERCEPTION;

			ItemData itemData = {0};
			BOOL comctl32600 = RunTimeHelper::IsCommCtrl6();
			if(!comctl32600) {
				itemData.lParam = pDetails->lParam;
				pDetails->lParam = GetNewItemID();
				pDetails->mask |= LVIF_PARAM;
			}

			#ifdef INCLUDESHELLBROWSERINTERFACE
				BOOL isShellItem = ((pDetails->mask & LVIF_ISASHELLITEM) == LVIF_ISASHELLITEM);
				pDetails->mask &= ~LVIF_ISASHELLITEM;
			#endif

			insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
			if(insertedItem != -1) {
				if(!comctl32600) {
					#ifdef USE_STL
						if(insertedItem >= static_cast<int>(itemIDs.size())) {
							itemIDs.push_back(static_cast<LONG>(pDetails->lParam));
						} else {
							itemIDs.insert(itemIDs.begin() + insertedItem, static_cast<LONG>(pDetails->lParam));
						}
						std::list<ItemData>::iterator iter2 = itemParams.begin();
						if(iter2 != itemParams.end()) {
							std::advance(iter2, insertedItem);
						}
						itemParams.insert(iter2, itemData);
					#else
						if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
							itemIDs.Add(static_cast<LONG>(pDetails->lParam));
						} else {
							itemIDs.InsertAt(insertedItem, static_cast<LONG>(pDetails->lParam));
						}
						POSITION p = itemParams.FindIndex(insertedItem);
						if(p) {
							itemParams.InsertBefore(p, itemData);
						} else {
							itemParams.AddTail(itemData);
						}
					#endif
				}

				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(shellBrowserInterface.pInternalMessageListener) {
						if(isShellItem) {
							pDetails->mask |= LVIF_ISASHELLITEM;
						}
						shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_INSERTEDITEM, ItemIndexToID(insertedItem), lParam);
					}
				#endif

				LVITEMINDEX insertedItemIndex = {insertedItem, 0};
				CComPtr<IListViewItem> pLvwItemItf = ClassFactory::InitListItem(insertedItemIndex, this);
				if(pLvwItemItf) {
					Raise_InsertedItem(pLvwItemItf);
				}
			}
		}
	} else {
		// finally pass the message to the listview
		LPLVITEM pDetails = reinterpret_cast<LPLVITEM>(lParam);
		pDetails->mask &= ~LVIF_NOINTERCEPTION;

		ItemData itemData = {0};
		BOOL comctl32600 = RunTimeHelper::IsCommCtrl6();
		if(!comctl32600) {
			itemData.lParam = pDetails->lParam;
			pDetails->lParam = GetNewItemID();
			pDetails->mask |= LVIF_PARAM;
		}

		#ifdef INCLUDESHELLBROWSERINTERFACE
			BOOL isShellItem = ((pDetails->mask & LVIF_ISASHELLITEM) == LVIF_ISASHELLITEM);
			pDetails->mask &= ~LVIF_ISASHELLITEM;
		#endif

		insertedItem = static_cast<int>(DefWindowProc(message, wParam, lParam));
		if(insertedItem != -1) {
			if(!comctl32600) {
				#ifdef USE_STL
					if(insertedItem >= static_cast<int>(itemIDs.size())) {
						itemIDs.push_back(static_cast<LONG>(pDetails->lParam));
					} else {
						itemIDs.insert(itemIDs.begin() + insertedItem, static_cast<LONG>(pDetails->lParam));
					}
					std::list<ItemData>::iterator iter2 = itemParams.begin();
					if(iter2 != itemParams.end()) {
						std::advance(iter2, insertedItem);
					}
					itemParams.insert(iter2, itemData);
				#else
					if(insertedItem >= static_cast<int>(itemIDs.GetCount())) {
						itemIDs.Add(static_cast<LONG>(pDetails->lParam));
					} else {
						itemIDs.InsertAt(insertedItem, static_cast<LONG>(pDetails->lParam));
					}
					POSITION p = itemParams.FindIndex(insertedItem);
					if(p) {
						itemParams.InsertBefore(p, itemData);
					} else {
						itemParams.AddTail(itemData);
					}
				#endif
			}

			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(shellBrowserInterface.pInternalMessageListener) {
					if(isShellItem) {
						pDetails->mask |= LVIF_ISASHELLITEM;
					}
					shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_INSERTEDITEM, ItemIndexToID(insertedItem), lParam);
				}
			#endif
		}
	}

	if(!(properties.disabledEvents & deListMouseEvents)) {
		if(insertedItem != -1) {
			// maybe we have a new (sub-)item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT h = 0;
			LVITEMINDEX newItemUnderMouse;
			int newSubItemUnderMouse = -1;
			HitTest(mousePosition.x, mousePosition.y, &h, &newItemUnderMouse, &newSubItemUnderMouse);
			HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
			if(newItemUnderMouse.iItem != itemUnderMouse.iItem || newItemUnderMouse.iGroup != itemUnderMouse.iGroup) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse.iItem != -1) {
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					Raise_ItemMouseLeave(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				subItemUnderMouse = newSubItemUnderMouse;

				if(itemUnderMouse.iItem != -1) {
					hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseEnter(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return insertedItem;
}

LRESULT ExplorerListView::OnRemoveAllGroups(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deGroupDeletionEvents)) {
		Raise_RemovingGroup(NULL, &cancel);
	}
	if(cancel == VARIANT_FALSE) {
		// finally pass the message to the listview
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			#ifdef USE_STL
				groups.clear();
			#else
				groups.RemoveAll();
			#endif
			if(!(properties.disabledEvents & deGroupDeletionEvents)) {
				Raise_RemovedGroup(NULL);
			}

			if(!(properties.disabledEvents & deListMouseEvents)) {
				// maybe we have a new (sub-)item below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);

				UINT h = 0;
				LVITEMINDEX newItemUnderMouse;
				int newSubItemUnderMouse = -1;
				HitTest(mousePosition.x, mousePosition.y, &h, &newItemUnderMouse, &newSubItemUnderMouse);
				HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
				if(newItemUnderMouse.iItem != itemUnderMouse.iItem || newItemUnderMouse.iGroup != itemUnderMouse.iGroup) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(itemUnderMouse.iItem != -1) {
						CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
						if(subItemUnderMouse > 0) {
							CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
							Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						}
						Raise_ItemMouseLeave(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}

					itemUnderMouse = newItemUnderMouse;
					subItemUnderMouse = newSubItemUnderMouse;

					if(itemUnderMouse.iItem != -1) {
						hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
						CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
						Raise_ItemMouseEnter(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						if(subItemUnderMouse > 0) {
							CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
							Raise_SubItemMouseEnter(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
						}
					}
				}
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnRemoveGroup(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IListViewGroup> pLvwGroupItf = NULL;
	CComObject<ListViewGroup>* pLvwGroupObj = NULL;
	if(!(properties.disabledEvents & deGroupDeletionEvents)) {
		CComObject<ListViewGroup>::CreateInstance(&pLvwGroupObj);
		pLvwGroupObj->AddRef();
		pLvwGroupObj->SetOwner(this);
		pLvwGroupObj->Attach(static_cast<int>(wParam));
		pLvwGroupObj->QueryInterface(IID_IListViewGroup, reinterpret_cast<LPVOID*>(&pLvwGroupItf));
		pLvwGroupObj->Release();
		Raise_RemovingGroup(pLvwGroupItf, &cancel);
	}

	if(cancel == VARIANT_FALSE) {
		CComPtr<IVirtualListViewGroup> pVLvwGroupItf = NULL;
		if(!(properties.disabledEvents & deGroupDeletionEvents)) {
			CComObject<VirtualListViewGroup>* pVLvwGroupObj = NULL;
			CComObject<VirtualListViewGroup>::CreateInstance(&pVLvwGroupObj);
			pVLvwGroupObj->AddRef();
			pVLvwGroupObj->SetOwner(this);
			if(pLvwGroupObj) {
				pLvwGroupObj->SaveState(pVLvwGroupObj);
			}

			pVLvwGroupObj->QueryInterface(IID_IVirtualListViewGroup, reinterpret_cast<LPVOID*>(&pVLvwGroupItf));
			pVLvwGroupObj->Release();
		}

		// finally pass the message to the listview
		lr = DefWindowProc(message, wParam, lParam);
		// the return value is not used, so assume success
		// TODO: Is it really not used?
		#ifdef USE_STL
			std::vector<int>::iterator iter = std::find(groups.begin(), groups.end(), static_cast<int>(wParam));
			if(iter != groups.end()) {
				groups.erase(iter);
			}
		#else
			for(size_t i = 0; i < groups.GetCount(); ++i) {
				if(groups[i] == static_cast<int>(wParam)) {
					groups.RemoveAt(i);
					break;
				}
			}
		#endif
		if(!(properties.disabledEvents & deGroupDeletionEvents)) {
			Raise_RemovedGroup(pVLvwGroupItf);
		}

		if(!(properties.disabledEvents & deListMouseEvents)) {
			// maybe we have a new (sub-)item below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT h = 0;
			LVITEMINDEX newItemUnderMouse;
			int newSubItemUnderMouse = -1;
			HitTest(mousePosition.x, mousePosition.y, &h, &newItemUnderMouse, &newSubItemUnderMouse);
			HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
			if(newItemUnderMouse.iItem != itemUnderMouse.iItem || newItemUnderMouse.iGroup != itemUnderMouse.iGroup) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(itemUnderMouse.iItem != -1) {
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseLeave(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
					Raise_ItemMouseLeave(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				itemUnderMouse = newItemUnderMouse;
				subItemUnderMouse = newSubItemUnderMouse;

				if(itemUnderMouse.iItem != -1) {
					hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
					CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
					Raise_ItemMouseEnter(pLvwItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					if(subItemUnderMouse > 0) {
						CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
						Raise_SubItemMouseEnter(pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnSetBkImage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL dontSynchronize = static_cast<BOOL>(wParam);
	wParam = 0;

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	LPLVBKIMAGE pBkImage = reinterpret_cast<LPLVBKIMAGE>(lParam);
	#ifdef DEBUG
		if(message == LVM_SETBKIMAGEA) {
			ATLASSERT_POINTER(pBkImage, LVBKIMAGEA);
		} else {
			ATLASSERT_POINTER(pBkImage, LVBKIMAGEW);
		}
	#endif
	if(!pBkImage) {
		return lr;
	}
	flags.hCurrentBackgroundBitmap = pBkImage->hbm;

	if(!dontSynchronize && lr && !IsInDesignMode()) {
		HBITMAP h = NULL;
		if(SUCCEEDED(VariantChangeType(&properties.bkImage, &properties.bkImage, 0, VT_I4))) {
			h = static_cast<HBITMAP>(LongToHandle(properties.bkImage.lVal));
		}
		if(properties.ownsBkImageBitmap && h && (h != pBkImage->hbm) && GetObjectType(h) == OBJ_BITMAP) {
			DeleteObject(h);
		}
		VariantClear(&properties.bkImage);

		switch(pBkImage->ulFlags & LVBKIF_SOURCE_MASK) {
			case LVBKIF_SOURCE_NONE:
				if(pBkImage->ulFlags & LVBKIF_TYPE_WATERMARK) {
					properties.bkImage.lVal = HandleToLong(pBkImage->hbm);
					properties.bkImage.vt = VT_I4;
				}
				break;
			case LVBKIF_SOURCE_HBITMAP:
				properties.bkImage.lVal = HandleToLong(pBkImage->hbm);
				properties.bkImage.vt = VT_I4;
				break;
			case LVBKIF_SOURCE_URL:
				if(message == LVM_SETBKIMAGEA) {
					properties.bkImage.bstrVal = CComBSTR(reinterpret_cast<LPLVBKIMAGEA>(pBkImage)->pszImage);
				} else {
					properties.bkImage.bstrVal = W2BSTR(reinterpret_cast<LPLVBKIMAGEW>(pBkImage)->pszImage);
				}
				properties.bkImage.vt = VT_BSTR;
				break;
		}

		properties.bkImagePositionX = pBkImage->xOffsetPercent;
		properties.bkImagePositionY = pBkImage->yOffsetPercent;
		properties.absoluteBkImagePosition = ((pBkImage->ulFlags & LVBKIF_FLAG_TILEOFFSET) == LVBKIF_FLAG_TILEOFFSET);

		if(pBkImage->ulFlags & LVBKIF_TYPE_WATERMARK) {
			properties.bkImageStyle = bisWatermark;
		} else {
			switch(pBkImage->ulFlags & LVBKIF_STYLE_MASK) {
				case LVBKIF_STYLE_NORMAL:
					properties.bkImageStyle = bisNormal;
					break;
				case LVBKIF_STYLE_TILE:
					properties.bkImageStyle = bisTiled;
					break;
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnSetExtendedListViewStyle(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if((wParam == 0) || ((wParam & LVS_EX_TRACKSELECT) == LVS_EX_TRACKSELECT)) {
		UINT b = cachedSettings.hotTracking;
		cachedSettings.hotTracking = ((lParam & LVS_EX_TRACKSELECT) == LVS_EX_TRACKSELECT);
		if(cachedSettings.hotTracking && !b) {
			// SysListView32 will start to track the mouse itself and we don't want to interfere
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnSetGroupInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = -1;
	PLVGROUP pGroup = reinterpret_cast<PLVGROUP>(lParam);

	if(((pGroup->mask & (LVGF_HEADER | LVGF_FOOTER)) != 0) && (static_cast<int>(wParam) == pGroup->iGroupId) && !IsComctl32Version610OrNewer()) {
		/* The sender wants to update the group's text. This isn't possible without changing the group's ID.
			  So we're kind and change the ID automatically. It will be reset below. */
		LVGROUP group = {0};
		group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
		int buffer = nextGroupID;
		group.iGroupId = GetUniqueGroupID();
		group.mask = LVGF_GROUPID;
		DefWindowProc(LVM_SETGROUPINFO, wParam, reinterpret_cast<LPARAM>(&group));

		pGroup->mask |= LVGF_GROUPID;
		// let's hope this message doesn't fail
		DefWindowProc(message, group.iGroupId, lParam);
		nextGroupID = buffer;
		// LVM_SETGROUPINFO returns the *old* group ID, so we've to set the return value ourselves
		lr = wParam;
	} else {
		lr = DefWindowProc(message, wParam, lParam);
		if((lr != -1) && (static_cast<int>(wParam) != pGroup->iGroupId)) {
			BOOL idWasChanged = FALSE;
			if(IsComctl32Version610OrNewer()) {
				idWasChanged = ((pGroup->mask & LVGF_GROUPID) == LVGF_GROUPID);
			} else {
				// Windows XP ignores the LVGROUP::mask member
				idWasChanged = TRUE;
			}

			if(idWasChanged) {
				// group id was changed
				ReplaceGroupID(static_cast<int>(wParam), pGroup->iGroupId);
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnSetHotCursor(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_EXLVW_HOTMOUSEICON) == S_FALSE || FireOnRequestEdit(DISPID_EXLVW_HOTMOUSEPOINTER) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.hotMouseIcon.dontGetPictureObject) {
		/* This message wasn't sent by ourselves, so we have to retrieve the new hotMouseIcon.pPictureDisp
		   object. */
		properties.hotMouseIcon.StopWatching();
		HCURSOR hCursor = reinterpret_cast<HCURSOR>(lParam);
		if(hCursor) {
			PICTDESC picture = {0};
			picture.cbSizeofstruct = sizeof(picture);
			picture.icon.hicon = static_cast<HICON>(hCursor);
			picture.picType = PICTYPE_ICON;

			// now create an IDispatch object out of the bitmap
			OleCreatePictureIndirect(&picture, IID_IDispatch, FALSE, reinterpret_cast<LPVOID*>(&properties.hotMouseIcon.pPictureDisp));
			properties.hotMousePointer = mpCustom;
		} else {
			properties.hotMouseIcon.pPictureDisp = NULL;
			properties.hotMousePointer = mpDefault;
		}
		properties.hotMouseIcon.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_EXLVW_HOTMOUSEICON);
		FireOnChanged(DISPID_EXLVW_HOTMOUSEPOINTER);
	} else {
		properties.hotMouseIcon.dontGetPictureObject = FALSE;
	}

	return lr;
}

LRESULT ExplorerListView::OnSetHotItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IListViewItem> pPreviousHotItem = NULL;
	CComPtr<IListViewItem> pNewHotItem = NULL;
	if(static_cast<int>(wParam) != cachedSettings.hotItem.iItem && !(properties.disabledEvents & deHotItemChangeEvents)) {
		pPreviousHotItem = ClassFactory::InitListItem(cachedSettings.hotItem, this);
		LVITEMINDEX itemIndex = {static_cast<int>(wParam), 0};
		pNewHotItem = ClassFactory::InitListItem(itemIndex, this);
		Raise_HotItemChanging(pPreviousHotItem, pNewHotItem, &cancel);
	}

	LRESULT lr = cachedSettings.hotItem.iItem;
	if(cancel == VARIANT_FALSE) {
		lr = DefWindowProc(message, wParam, lParam);

		if(static_cast<int>(wParam) != cachedSettings.hotItem.iItem) {
			cachedSettings.hotItem.iItem = static_cast<int>(wParam);
			cachedSettings.hotItem.iGroup = 0;
			if(!(properties.disabledEvents & deHotItemChangeEvents)) {
				Raise_HotItemChanged(pPreviousHotItem, pNewHotItem);
			}
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL useViewModeHack = FALSE;

	HIMAGELIST buffer = NULL;
	/* We must store the new settings before the call to DefWindowProc, because we may need them while
	   handling the custom draw notifications this message probably will cause. */
	switch(wParam) {
		case LVSIL_FOOTERITEMS:
			cachedSettings.hFooterItemsImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;
		case LVSIL_GROUPHEADER:
			cachedSettings.hGroupsImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;
		case LVSIL_NORMAL:
			if(IsInViewMode(vTiles, FALSE)) {
				useViewModeHack = (!cachedSettings.hExtraLargeImageList && lParam);
				cachedSettings.hExtraLargeImageList = reinterpret_cast<HIMAGELIST>(lParam);
			} else {
				cachedSettings.hLargeImageList = reinterpret_cast<HIMAGELIST>(lParam);
			}
			break;
		case LVSIL_SMALL:
			// SysListView32 will overwrite our settings
			buffer = reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_NORMAL, 0));
			cachedSettings.hSmallImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;
		case LVSIL_STATE:
			cachedSettings.hStateImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// TODO: Isn't there a better way for correcting the values if an error occured?
	switch(wParam) {
		case LVSIL_FOOTERITEMS:
			cachedSettings.hFooterItemsImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_FOOTERITEMS, 0));
			break;
		case LVSIL_GROUPHEADER:
			cachedSettings.hGroupsImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_GROUPHEADER, 0));
			break;
		case LVSIL_NORMAL:
			if(IsInViewMode(vTiles, FALSE)) {
				cachedSettings.hExtraLargeImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_NORMAL, 0));
				useViewModeHack = useViewModeHack && (cachedSettings.hExtraLargeImageList == reinterpret_cast<HIMAGELIST>(lParam));
			} else {
				cachedSettings.hLargeImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_NORMAL, 0));
			}
			break;
		case LVSIL_SMALL:
			cachedSettings.hSmallImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_SMALL, 0));
			// SysListView32 has overwritten our settings
			containedSysHeader32.SendMessage(HDM_SETIMAGELIST, HDSIL_NORMAL, reinterpret_cast<LPARAM>(buffer));
			break;
		case LVSIL_STATE:
			cachedSettings.hStateImageList = reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_STATE, 0));
			break;
	}

	if(useViewModeHack) {
		ViewConstants view;
		if(SUCCEEDED(get_View(&view))) {
			switch(view) {
				case vIcons:
					put_View(vList);
					break;
				case vSmallIcons:
					put_View(vList);
					break;
				case vList:
					put_View(vSmallIcons);
					break;
				case vDetails:
					put_View(vList);
					break;
				case vTiles:
					put_View(vList);
					break;
				case vExtendedTiles:
					put_View(vList);
					break;
			}
			put_View(view);
		}
	}

	return lr;
}

LRESULT ExplorerListView::OnSetInfoTip(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	PLVSETINFOTIP pInfoTip = reinterpret_cast<PLVSETINFOTIP>(lParam);
	ATLASSERT_POINTER(pInfoTip, LVSETINFOTIP);
	if(!pInfoTip) {
		wasHandled = FALSE;
		return FALSE;
	}

	LVITEMINDEX itemIndex = {pInfoTip->iItem, 0};
	CComPtr<IListViewItem> pListItem = ClassFactory::InitListItem(itemIndex, this);
	// when you change this code, be careful with who frees which string
	CComBSTR infoTipText = L"";
	infoTipText.Append(pInfoTip->pszText);
	VARIANT_BOOL cancel = VARIANT_FALSE;

	Raise_SettingItemInfoTipText(pListItem, &infoTipText, &cancel);

	LRESULT lr = FALSE;
	if(cancel == VARIANT_FALSE) {
		LPWSTR p = pInfoTip->pszText;
		pInfoTip->pszText = OLE2W(infoTipText);
		lr = DefWindowProc(message, wParam, lParam);
		pInfoTip->pszText = p;
	}
	return lr;
}

LRESULT ExplorerListView::OnSetItem(UINT message, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPLVITEM pItemData = reinterpret_cast<LPLVITEM>(lParam);
	#ifdef DEBUG
		if(message == LVM_SETITEMA) {
			ATLASSERT_POINTER(pItemData, LVITEMA);
		} else {
			ATLASSERT_POINTER(pItemData, LVITEMW);
		}
	#endif
	if(!pItemData) {
		wasHandled = FALSE;
		return TRUE;
	}

	if(pItemData->mask & LVIF_NOINTERCEPTION) {
		// SysListView32 shouldn't see this flag
		pItemData->mask &= ~LVIF_NOINTERCEPTION;
	} else if(pItemData->mask == LVIF_PARAM) {
		if(!RunTimeHelper::IsCommCtrl6()) {
			// simply update the 'itemParams' list
			#ifdef USE_STL
				std::list<ItemData>::iterator iter = itemParams.begin();
				if(iter != itemParams.end()) {
					std::advance(iter, pItemData->iItem);
					if(iter != itemParams.end()) {
						iter->lParam = pItemData->lParam;
						return TRUE;
					}
				}
			#else
				POSITION p = itemParams.FindIndex(pItemData->iItem);
				if(p) {
					itemParams.GetAt(p).lParam = pItemData->lParam;
					return TRUE;
				}
			#endif
			// item not found
			return FALSE;
		}
	} else if(pItemData->mask & LVIF_PARAM) {
		if(!RunTimeHelper::IsCommCtrl6()) {
			#ifdef USE_STL
				std::list<ItemData>::iterator iter = itemParams.begin();
				if(iter != itemParams.end()) {
					std::advance(iter, pItemData->iItem);
					if(iter != itemParams.end()) {
						iter->lParam = pItemData->lParam;
						pItemData->mask &= ~LVIF_PARAM;
					}
				}
			#else
				POSITION p = itemParams.FindIndex(pItemData->iItem);
				if(p) {
					itemParams.GetAt(p).lParam = pItemData->lParam;
					pItemData->mask &= ~LVIF_PARAM;
				}
			#endif
		}
	}

	// let SysListView32 handle the message
	wasHandled = FALSE;
	return TRUE;
}

LRESULT ExplorerListView::OnSetItemCount(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// forward the message
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.virtualItemCount = static_cast<long>(wParam);
	}

	return lr;
}

LRESULT ExplorerListView::OnSetView(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL resetHeaderChevron = FALSE;
	if(containedSysHeader32.IsWindow()) {
		resetHeaderChevron = ((containedSysHeader32.GetStyle() & HDS_OVERFLOW) == HDS_OVERFLOW);
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr == 1) {
		if(wParam == LV_VIEW_TILE) {
			// set the extra large imagelist
			SendMessage(LVM_SETIMAGELIST, LVSIL_NORMAL, reinterpret_cast<LPARAM>(cachedSettings.hExtraLargeImageList));
		} else {
			// set the large imagelist
			SendMessage(LVM_SETIMAGELIST, LVSIL_NORMAL, reinterpret_cast<LPARAM>(cachedSettings.hLargeImageList));
		}
		if(resetHeaderChevron) {
			// LVM_SETVIEW removes the HDS_OVERFLOW style
			containedSysHeader32.ModifyStyle(0, HDS_OVERFLOW);
		}

		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pInternalMessageListener) {
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_CHANGEDVIEW, wParam, 0);
			}
		#endif
	}

	return lr;
}

LRESULT ExplorerListView::OnSetWorkAreas(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = 0;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IVirtualListViewWorkAreas> pVLvwWorkAreasItf = NULL;
	CComObject<VirtualListViewWorkAreas>* pVLvwWorkAreasObj = NULL;
	CComObject<VirtualListViewWorkAreas>::CreateInstance(&pVLvwWorkAreasObj);
	pVLvwWorkAreasObj->AddRef();
	pVLvwWorkAreasObj->Attach(static_cast<int>(wParam), reinterpret_cast<LPRECT>(lParam));
	pVLvwWorkAreasObj->QueryInterface(IID_IVirtualListViewWorkAreas, reinterpret_cast<LPVOID*>(&pVLvwWorkAreasItf));
	pVLvwWorkAreasObj->Release();
	Raise_ChangingWorkAreas(pVLvwWorkAreasItf, &cancel);

	if(cancel == VARIANT_FALSE) {
		lr = DefWindowProc(message, wParam, lParam);

		CComPtr<IListViewWorkAreas> pLvwWorkAreasItf = NULL;
		get_WorkAreas(&pLvwWorkAreasItf);
		Raise_ChangedWorkAreas(pLvwWorkAreasItf);
	}

	return lr;
}

LRESULT ExplorerListView::OnSortGroups(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SortGroupsDetails newLParam = {0};
	newLParam.pRealComparisonFunction = reinterpret_cast<PFNLVGROUPCOMPARE>(wParam);
	newLParam.pAppDefinedInformation = reinterpret_cast<LPVOID>(lParam);
	newLParam.pGroups = &groups;

	return DefWindowProc(message, reinterpret_cast<WPARAM>(::MyLVGroupCompare), reinterpret_cast<LPARAM>(&newLParam));
}

LRESULT ExplorerListView::OnSortItems(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	#ifdef USE_STL
		size_t itemCount = itemIDs.size();
	#else
		size_t itemCount = itemIDs.GetCount();
	#endif
	if(itemCount > 1) {
		// we've to synchronize 'itemIDs' and 'itemParams'

		SortItemsDetails newWParam = {0};
		newWParam.pRealComparisonFunction = reinterpret_cast<PFNLVCOMPARE>(lParam);
		newWParam.pAppDefinedInformation = wParam;

		PLONG pNewItemOrder = new LONG[itemCount];
		// store current item order
		for(size_t i = 0; i < itemCount; ++i) {
			pNewItemOrder[i] = itemIDs[i];
		}

		// sort
		qsort_s(pNewItemOrder, itemCount, sizeof(LONG), ::MyQuickSortCompareFunc, &newWParam);

		// synchronize 'itemIDs' and 'itemParams'
		#ifdef USE_STL
			std::vector<LONG> itemIDsBackup(itemIDs.begin(), itemIDs.end());
			std::list<ItemData> itemParamsBackup(itemParams.begin(), itemParams.end());
			itemIDs.clear();
			itemParams.clear();
			for(int newIndex = 0; newIndex < static_cast<int>(itemCount); ++newIndex) {
				int oldIndex = -1;
				std::vector<LONG>::iterator iter = std::find(itemIDsBackup.begin(), itemIDsBackup.end(), pNewItemOrder[newIndex]);
				if(iter != itemIDsBackup.end()) {
					oldIndex = std::distance(itemIDsBackup.begin(), iter);
				}
				itemIDs.push_back(itemIDsBackup[oldIndex]);
				std::list<ItemData>::iterator iter2 = itemParamsBackup.begin();
				std::advance(iter2, oldIndex);
				itemParams.insert(itemParams.end(), *iter2);
			}
		#else
			CAtlArray<LONG> itemIDsBackup;
			itemIDsBackup.Append(itemIDs);
			CAtlList<ItemData> itemParamsBackup;
			itemParamsBackup.AddTailList(&itemParams);
			itemIDs.RemoveAll();
			itemParams.RemoveAll();
			for(int newIndex = 0; newIndex < static_cast<int>(itemCount); ++newIndex) {
				int oldIndex = -1;
				for(size_t i = 0; i < itemIDsBackup.GetCount(); ++i) {
					if(itemIDsBackup[i] == pNewItemOrder[newIndex]) {
						oldIndex = i;
						break;
					}
				}
				itemIDs.Add(itemIDsBackup[oldIndex]);
				POSITION p = itemParamsBackup.FindIndex(oldIndex);
				ATLASSERT(p);
				itemParams.AddTail(itemParamsBackup.GetAt(p));
			}
		#endif

		// now forward the message and sort the items based on 'pNewItemOrder'
		newWParam.pNewItemIDs = &itemIDs;
		LRESULT lr = DefWindowProc(message, reinterpret_cast<WPARAM>(&newWParam), reinterpret_cast<LPARAM>(::MyLVCompareFunc));
		if(!lr) {
			// rollback
			#ifdef USE_STL
				itemParams.clear();
				itemParams.resize(itemParamsBackup.size());
				std::copy(itemParamsBackup.begin(), itemParamsBackup.end(), itemParams.begin());
				itemIDs.clear();
				itemIDs.resize(itemIDsBackup.size());
				std::copy(itemIDsBackup.begin(), itemIDsBackup.end(), itemIDs.begin());
			#else
				itemParams.RemoveAll();
				itemParams.AddTailList(&itemParamsBackup);
				itemIDs.RemoveAll();
				itemIDs.Append(itemIDsBackup);
			#endif
		}

		// cleanup
		delete[] pNewItemOrder;
		return TRUE;
	}
	wasHandled = FALSE;
	return TRUE;
}

LRESULT ExplorerListView::OnSortItemsEx(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	#ifdef USE_STL
		size_t itemCount = itemIDs.size();
	#else
		size_t itemCount = itemIDs.GetCount();
	#endif
	if(itemCount > 1) {
		// we've to synchronize 'itemIDs' and 'itemParams'

		SortItemsDetailsEx newWParam = {0};
		newWParam.pRealComparisonFunction = reinterpret_cast<PFNLVCOMPARE>(lParam);
		newWParam.pAppDefinedInformation = wParam;

		PINT pNewItemOrder = new int[itemCount];
		// store current item order
		for(int i = 0; i < static_cast<int>(itemCount); ++i) {
			pNewItemOrder[i] = i;
		}

		// sort
		qsort_s(pNewItemOrder, itemCount, sizeof(int), ::MyQuickSortCompareFuncEx, &newWParam);
		// build a map that maps old indexes to new indexes
		#ifdef USE_STL
			std::unordered_map< int, int, ItemIndexHasher<int> > newItemIndexes;
		#else
			CAtlMap<int, int> newItemIndexes;
		#endif
		for(int newIndex = 0; newIndex < static_cast<int>(itemCount); ++newIndex) {
			newItemIndexes[pNewItemOrder[newIndex]] = newIndex;
		}

		// synchronize 'itemIDs' and 'itemParams'
		#ifdef USE_STL
			std::vector<LONG> itemIDsBackup(itemIDs.begin(), itemIDs.end());
			std::list<ItemData> itemParamsBackup(itemParams.begin(), itemParams.end());
			itemIDs.clear();
			itemParams.clear();
			for(int newIndex = 0; newIndex < static_cast<int>(itemCount); ++newIndex) {
				int oldIndex = pNewItemOrder[newIndex];
				itemIDs.push_back(itemIDsBackup[oldIndex]);
				std::list<ItemData>::iterator iter = itemParamsBackup.begin();
				std::advance(iter, oldIndex);
				itemParams.insert(itemParams.end(), *iter);
			}
		#else
			CAtlArray<LONG> itemIDsBackup;
			itemIDsBackup.Append(itemIDs);
			CAtlList<ItemData> itemParamsBackup;
			itemParamsBackup.AddTailList(&itemParams);
			itemIDs.RemoveAll();
			itemParams.RemoveAll();
			for(int newIndex = 0; newIndex < static_cast<int>(itemCount); ++newIndex) {
				int oldIndex = pNewItemOrder[newIndex];
				itemIDs.Add(itemIDsBackup[oldIndex]);
				POSITION p = itemParamsBackup.FindIndex(oldIndex);
				ATLASSERT(p);
				itemParams.AddTail(itemParamsBackup.GetAt(p));
			}
		#endif

		// now forward the message and sort the items based on 'newItemIndexes'
		newWParam.pNewItemIndexes = &newItemIndexes;
		LRESULT lr = DefWindowProc(message, reinterpret_cast<WPARAM>(&newWParam), reinterpret_cast<LPARAM>(::MyLVCompareFuncEx));
		if(!lr) {
			// rollback
			#ifdef USE_STL
				itemParams.clear();
				itemParams.resize(itemParamsBackup.size());
				std::copy(itemParamsBackup.begin(), itemParamsBackup.end(), itemParams.begin());
				itemIDs.clear();
				itemIDs.resize(itemIDsBackup.size());
				std::copy(itemIDsBackup.begin(), itemIDsBackup.end(), itemIDs.begin());
			#else
				itemParams.RemoveAll();
				itemParams.AddTailList(&itemParamsBackup);
				itemIDs.RemoveAll();
				itemIDs.Append(itemIDsBackup);
			#endif
		}

		// cleanup
		delete[] pNewItemOrder;
		return TRUE;
	}
	wasHandled = FALSE;
	return TRUE;
}

LRESULT ExplorerListView::OnEditChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_EditKeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ExplorerListView::OnEditContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		containedEdit.ScreenToClient(&mousePosition);
	}
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_EditContextMenu(button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);
	if(showDefaultMenu != VARIANT_FALSE) {
		wasHandled = FALSE;
	}

	return 0;
}

LRESULT ExplorerListView::OnEditDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedEditControlWindow(containedEdit);
	return 0;
}

LRESULT ExplorerListView::OnEditKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deEditKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_EditKeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnEditKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deEditKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_EditKeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedEdit.DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnEditKillFocus(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_EditLostFocus();
	return 0;
}

LRESULT ExplorerListView::OnEditLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_EditDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_EditMDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_EditMouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_EditMouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y);
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_EditMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		containedEdit.ScreenToClient(&mousePosition);
		Raise_EditMouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_EditRDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	LRESULT lr = CComControl<ExplorerListView>::OnSetFocus(message, wParam, lParam, wasHandled);

	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(containedEdit.m_hWnd);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode_Edit();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(containedEdit.m_hWnd, ime);
			}
		}
	}

	Raise_EditGotFocus();
	return lr;
}

LRESULT ExplorerListView::OnEditXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_EditXDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_EditMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnEditXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_EditMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_HeaderKeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ExplorerListView::OnHeaderContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		containedSysHeader32.ScreenToClient(&mousePosition);
	}
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_HeaderContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT ExplorerListView::OnHeaderDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_DestroyedHeaderControlWindow(containedSysHeader32);
	return 0;
}

LRESULT ExplorerListView::OnHeaderKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deHeaderKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_HeaderKeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedSysHeader32.DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnHeaderKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deHeaderKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_HeaderKeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return containedSysHeader32.DefWindowProc(message, wParam, lParam);
}

LRESULT ExplorerListView::OnHeaderKillFocus(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_HeaderLostFocus();
	return 0;
}

LRESULT ExplorerListView::OnHeaderLButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* Do this to ensure that the HDN_ITEMDBLCLICK notification is sent before we raise the HeaderDblClick
	   event. */
	flags.raisedHeaderDblClick = FALSE;
	LRESULT lr = containedSysHeader32.DefWindowProc(message, wParam, lParam);

	if(!(properties.disabledEvents & deHeaderClickEvents)) {
		if(!flags.raisedHeaderDblClick) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(wParam, button, shift);
			button = 1/*MouseButtonConstants.vbLeftButton*/;
			Raise_HeaderDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
		}
	}
	flags.raisedHeaderDblClick = FALSE;

	return lr;
}

LRESULT ExplorerListView::OnHeaderLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_HeaderMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ExplorerListView::OnHeaderLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// do this to ensure that the HDN_ITEMCLICK notification is sent before we raise the HeaderMouseUp event
	LRESULT lr = containedSysHeader32.DefWindowProc(message, wParam, lParam);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_HeaderMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return lr;
}

LRESULT ExplorerListView::OnHeaderMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_HeaderMDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_HeaderMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ExplorerListView::OnHeaderMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_HeaderMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ExplorerListView::OnHeaderMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deHeaderMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_HeaderMouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_HeaderMouseLeave(button, shift, mouseStatus_Header.lastPosition.x, mouseStatus_Header.lastPosition.y);
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_HeaderMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus_Header.enteredControl) {
		mouseStatus_Header.EnterControl();
	}

	if(dragDropStatus.HeaderIsDragging()) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		containedSysHeader32.ClientToScreen(&mousePosition);
		ScreenToClient(&mousePosition);
		Raise_HeaderDragMouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), mousePosition.x, mousePosition.y);
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderMouseWheel(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		if(message == WM_MOUSEHWHEEL) {
			// wParam and lParam seem to be 0
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			GetCursorPos(&mousePosition);
		} else {
			WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		}
		containedSysHeader32.ScreenToClient(&mousePosition);
		Raise_HeaderMouseWheel(button, shift, mousePosition.x, mousePosition.y, message == WM_MOUSEHWHEEL ? saHorizontal : saVertical, GET_WHEEL_DELTA_WPARAM(wParam));
	} else if(!mouseStatus_Header.enteredControl) {
		mouseStatus_Header.EnterControl();
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_HeaderRDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_HeaderMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ExplorerListView::OnHeaderRButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_HeaderMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	/* SysHeader32 sends a NM_RCLICK to the listview. This notification will raise listview events, so
	   we set a flag to ignore it. */
	flags.ignoreRClickNotification = TRUE;
	LRESULT lr = containedSysHeader32.DefWindowProc(message, wParam, lParam);
	flags.ignoreRClickNotification = FALSE;

	return lr;
}

LRESULT ExplorerListView::OnHeaderSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	LRESULT lr = CComControl<ExplorerListView>::OnSetFocus(message, wParam, lParam, wasHandled);
	Raise_HeaderGotFocus();
	return lr;
}

LRESULT ExplorerListView::OnHeaderXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deHeaderClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_HeaderXDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_HeaderMouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ExplorerListView::OnHeaderXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_HeaderMouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ExplorerListView::OnHeaderGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.draggedColumn != -1) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(dragDropStatus.draggedColumn, this, FALSE);
		if(flags.usingThemes && properties.headerOLEDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = CreateVistaOLEHeaderDragImage(pLvwColumn, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEHeaderDragImage(pLvwColumn, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = reinterpret_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT ExplorerListView::OnHeaderGetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	BOOL overwriteLParam = FALSE;

	LPHDITEM pColumnData = reinterpret_cast<LPHDITEM>(lParam);
	#ifdef DEBUG
		if(message == HDM_GETITEMA) {
			ATLASSERT_POINTER(pColumnData, HDITEMA);
		} else {
			ATLASSERT_POINTER(pColumnData, HDITEMW);
		}
	#endif
	if(!pColumnData) {
		return DefWindowProc(message, wParam, lParam);
	}

	if(pColumnData->mask & HDI_NOINTERCEPTION) {
		// SysHeader32 shouldn't see this flag
		pColumnData->mask &= ~HDI_NOINTERCEPTION;
	} else if(pColumnData->mask == HDI_LPARAM) {
		// just look up in the 'columnParams' list
		#ifdef USE_STL
			std::list<ColumnData>::iterator iter = columnParams.begin();
			if(iter != columnParams.end()) {
				std::advance(iter, wParam);
				if(iter != columnParams.end()) {
					pColumnData->lParam = iter->lParam;
					return TRUE;
				}
			}
		#else
			POSITION p = columnParams.FindIndex(wParam);
			if(p) {
				pColumnData->lParam = columnParams.GetAt(p).lParam;
				return TRUE;
			}
		#endif
		// column not found
		return FALSE;
	} else if(pColumnData->mask & HDI_LPARAM) {
		overwriteLParam = TRUE;
	}

	// forward the message
	LRESULT lr = containedSysHeader32.DefWindowProc(message, wParam, lParam);

	if(overwriteLParam) {
		#ifdef USE_STL
			std::list<ColumnData>::iterator iter = columnParams.begin();
			if(iter != columnParams.end()) {
				std::advance(iter, wParam);
				if(iter != columnParams.end()) {
					pColumnData->lParam = iter->lParam;
				}
			}
		#else
			POSITION p = columnParams.FindIndex(wParam);
			if(p) {
				pColumnData->lParam = columnParams.GetAt(p).lParam;
			}
		#endif
	}

	return lr;
}

LRESULT ExplorerListView::OnHeaderInsertItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	int insertedColumn = -1;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		BOOL notifyShellListView = ((bufferedColumnData.mask & LVCF_NOTIFYSHELLLISTVIEW) == LVCF_NOTIFYSHELLLISTVIEW);
		bufferedColumnData.mask &= ~LVCF_NOTIFYSHELLLISTVIEW;
	#endif

	if(!(properties.disabledEvents & deColumnInsertionEvents)) {
		CComObject<VirtualListViewColumn>* pVLvwColumnObj = NULL;
		CComPtr<IVirtualListViewColumn> pVLvwColumnItf = NULL;
		CComObject<VirtualListViewColumn>::CreateInstance(&pVLvwColumnObj);
		pVLvwColumnObj->AddRef();
		pVLvwColumnObj->SetOwner(this);

		#ifdef UNICODE
			BOOL mustConvert = (message == HDM_INSERTITEMA);
		#else
			BOOL mustConvert = (message == HDM_INSERTITEMW);
		#endif
		if(mustConvert) {
			#ifdef UNICODE
				LPHDITEMA pColumnData = reinterpret_cast<LPHDITEMA>(lParam);
				HDITEMW convertedColumnData = {0};
				LPSTR p = NULL;
				if(pColumnData->pszText != LPSTR_TEXTCALLBACKA) {
					p = pColumnData->pszText;
				}
				CA2W converter(p);
				if(pColumnData->pszText == LPSTR_TEXTCALLBACKA) {
					convertedColumnData.pszText = LPSTR_TEXTCALLBACKW;
				} else {
					convertedColumnData.pszText = converter;
				}
			#else
				LPHDITEMW pColumnData = reinterpret_cast<LPHDITEMW>(lParam);
				HDITEMA convertedColumnData = {0};
				LPWSTR p = NULL;
				if(pColumnData->pszText != LPSTR_TEXTCALLBACKW) {
					p = pColumnData->pszText;
				}
				CW2A converter(p);
				if(pColumnData->pszText == LPSTR_TEXTCALLBACKW) {
					convertedColumnData.pszText = LPSTR_TEXTCALLBACKA;
				} else {
					convertedColumnData.pszText = converter;
				}
			#endif
			convertedColumnData.cchTextMax = pColumnData->cchTextMax;
			convertedColumnData.mask = pColumnData->mask;
			convertedColumnData.cxy = pColumnData->cxy;
			convertedColumnData.hbm = pColumnData->hbm;
			convertedColumnData.fmt = pColumnData->fmt;
			convertedColumnData.iImage = pColumnData->iImage;
			convertedColumnData.lParam = pColumnData->lParam;
			convertedColumnData.iOrder = pColumnData->iOrder;
			// the HDITEM struct may end here, so be more careful with the remaining members
			if(convertedColumnData.mask & HDI_FILTER) {
				convertedColumnData.type = pColumnData->type;
				convertedColumnData.pvFilter = pColumnData->pvFilter;
			}
			pVLvwColumnObj->LoadState(&convertedColumnData, &bufferedColumnData);
		} else {
			pVLvwColumnObj->LoadState(reinterpret_cast<LPHDITEM>(lParam), &bufferedColumnData);
		}

		if(!(pVLvwColumnObj->properties.headerSettings.mask & HDI_ORDER)) {
			pVLvwColumnObj->properties.headerSettings.mask |= HDI_ORDER;
			pVLvwColumnObj->properties.headerSettings.iOrder = static_cast<int>(wParam);
		}
		pVLvwColumnObj->QueryInterface(IID_IVirtualListViewColumn, reinterpret_cast<LPVOID*>(&pVLvwColumnItf));
		pVLvwColumnObj->Release();

		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_InsertingColumn(pVLvwColumnItf, &cancel);
		pVLvwColumnObj = NULL;
		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the header
			LPHDITEM pDetails = reinterpret_cast<LPHDITEM>(lParam);
			pDetails->mask &= ~HDI_NOINTERCEPTION;

			ColumnData columnData = {0};
			columnData.lParam = pDetails->lParam;
			pDetails->lParam = GetNewColumnID();
			pDetails->mask |= HDI_LPARAM;
			columnData.localeID = static_cast<LCID>(-1);

			insertedColumn = static_cast<int>(containedSysHeader32.DefWindowProc(message, wParam, lParam));
			if(insertedColumn != -1) {
				#ifdef USE_STL
					for(std::unordered_map< LONG, int, ItemIndexHasher<LONG> >::iterator iter = columnIndexes.begin(); iter != columnIndexes.end(); ++iter) {
						if(iter->second >= insertedColumn) {
							++(iter->second);
						}
					}
					std::list<ColumnData>::iterator iter2 = columnParams.begin();
					if(iter2 != columnParams.end()) {
						std::advance(iter2, insertedColumn);
					}
					columnParams.insert(iter2, columnData);
				#else
					POSITION p = columnIndexes.GetStartPosition();
					while(p) {
						CAtlMap<LONG, int>::CPair* pPair = columnIndexes.GetAt(p);
						if(pPair->m_value >= insertedColumn) {
							++(pPair->m_value);
						}
						columnIndexes.GetNext(p);
					}
					p = columnParams.FindIndex(insertedColumn);
					if(p) {
						columnParams.InsertBefore(p, columnData);
					} else {
						columnParams.AddTail(columnData);
					}
				#endif
				columnIndexes[static_cast<const LONG>(pDetails->lParam)] = insertedColumn;

				#ifdef INCLUDESHELLBROWSERINTERFACE
					if(notifyShellListView) {
						if(shellBrowserInterface.pInternalMessageListener) {
							shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_INSERTEDCOLUMN, ColumnIndexToID(insertedColumn), 0);
						}
					}
				#endif

				CComPtr<IListViewColumn> pLvwColumnItf = ClassFactory::InitColumn(insertedColumn, this);
				if(pLvwColumnItf) {
					Raise_InsertedColumn(pLvwColumnItf);
				}
			}
		}
	} else {
		// finally pass the message to the header
		LPHDITEM pDetails = reinterpret_cast<LPHDITEM>(lParam);
		pDetails->mask &= ~HDI_NOINTERCEPTION;

		ColumnData columnData = {0};
		columnData.lParam = pDetails->lParam;
		pDetails->lParam = GetNewColumnID();
		pDetails->mask |= HDI_LPARAM;
		columnData.localeID = static_cast<LCID>(-1);

		insertedColumn = static_cast<int>(containedSysHeader32.DefWindowProc(message, wParam, lParam));
		if(insertedColumn != -1) {
			#ifdef USE_STL
				for(std::unordered_map< LONG, int, ItemIndexHasher<LONG> >::iterator iter = columnIndexes.begin(); iter != columnIndexes.end(); ++iter) {
					if(iter->second >= insertedColumn) {
						++(iter->second);
					}
				}
				std::list<ColumnData>::iterator iter2 = columnParams.begin();
				if(iter2 != columnParams.end()) {
					std::advance(iter2, insertedColumn);
				}
				columnParams.insert(iter2, columnData);
			#else
				POSITION p = columnIndexes.GetStartPosition();
				while(p) {
					CAtlMap<LONG, int>::CPair* pPair = columnIndexes.GetAt(p);
					if(pPair->m_value >= insertedColumn) {
						++(pPair->m_value);
					}
					columnIndexes.GetNext(p);
				}
				p = columnParams.FindIndex(insertedColumn);
				if(p) {
					columnParams.InsertBefore(p, columnData);
				} else {
					columnParams.AddTail(columnData);
				}
			#endif
			columnIndexes[static_cast<const LONG>(pDetails->lParam)] = insertedColumn;

			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(notifyShellListView) {
					if(shellBrowserInterface.pInternalMessageListener) {
						shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_INSERTEDCOLUMN, ColumnIndexToID(insertedColumn), 0);
					}
				}
			#endif
		}
	}

	if(!(properties.disabledEvents & deHeaderMouseEvents)) {
		if(insertedColumn != -1) {
			// maybe we have a new column below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			UINT hitTestDetails = 0;
			int newColumnUnderMouse = HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newColumnUnderMouse != columnUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(columnUnderMouse != -1) {
					CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnUnderMouse, this);
					Raise_ColumnMouseLeave(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
				}

				columnUnderMouse = newColumnUnderMouse;
				if(columnUnderMouse) {
					CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnUnderMouse, this);
					Raise_ColumnMouseEnter(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
				}
			}
		}
	}

	return insertedColumn;
}

LRESULT ExplorerListView::OnHeaderSetFilterChangeTimeout(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	properties.filterChangedTimeout = static_cast<long>(lParam);
	return 0;
}

LRESULT ExplorerListView::OnHeaderSetImageList(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	/* We must store the new settings before the call to DefWindowProc, because we may need them while
	   handling the custom draw notifications this message probably will cause. */
	switch(wParam) {
		/*case HDSIL_NORMAL:
			cachedSettings.hHeaderImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;*/
		case HDSIL_STATE:
			cachedSettings.hHeaderStateImageList = reinterpret_cast<HIMAGELIST>(lParam);
			break;
	}

	LRESULT lr = containedSysHeader32.DefWindowProc(message, wParam, lParam);

	// TODO: Isn't there a better way for correcting the values if an error occured?
	switch(wParam) {
		/*case HDSIL_NORMAL:
			cachedSettings.hHeaderImageList = reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_NORMAL, 0));
			break;*/
		case HDSIL_STATE:
			cachedSettings.hHeaderStateImageList = reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_STATE, 0));
			break;
	}

	return lr;
}

LRESULT ExplorerListView::OnHeaderSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LPHDITEM pColumnData = reinterpret_cast<LPHDITEM>(lParam);
	#ifdef DEBUG
		if(message == HDM_SETITEMA) {
			ATLASSERT_POINTER(pColumnData, HDITEMA);
		} else {
			ATLASSERT_POINTER(pColumnData, HDITEMW);
		}
	#endif
	if(!pColumnData) {
		wasHandled = FALSE;
		return TRUE;
	}

	if(pColumnData->mask & HDI_NOINTERCEPTION) {
		// SysHeader32 shouldn't see this flag
		pColumnData->mask &= ~HDI_NOINTERCEPTION;
	} else if(pColumnData->mask == HDI_LPARAM) {
		// simply update the 'columnParams' list
		#ifdef USE_STL
			std::list<ColumnData>::iterator iter = columnParams.begin();
			if(iter != columnParams.end()) {
				std::advance(iter, wParam);
				if(iter != columnParams.end()) {
					iter->lParam = pColumnData->lParam;
					return TRUE;
				}
			}
		#else
			POSITION p = columnParams.FindIndex(wParam);
			if(p) {
				columnParams.GetAt(p).lParam = pColumnData->lParam;
				return TRUE;
			}
		#endif
		// column not found
		return FALSE;
	} else if(pColumnData->mask & HDI_LPARAM) {
		#ifdef USE_STL
			std::list<ColumnData>::iterator iter = columnParams.begin();
			if(iter != columnParams.end()) {
				std::advance(iter, wParam);
				if(iter != columnParams.end()) {
					iter->lParam = pColumnData->lParam;
					pColumnData->mask &= ~HDI_LPARAM;
				}
			}
		#else
			POSITION p = columnParams.FindIndex(wParam);
			if(p) {
				columnParams.GetAt(p).lParam = pColumnData->lParam;
				pColumnData->mask &= ~LVIF_PARAM;
			}
		#endif
	}

	// let SysHeader32 handle the message
	wasHandled = FALSE;
	return TRUE;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	HRESULT ExplorerListView::OnInternalAddColumn(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXLVMADDCOLUMNDATA pDetails = reinterpret_cast<LPEXLVMADDCOLUMNDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXLVMADDCOLUMNDATA);
		if(!pDetails) {
			return E_POINTER;
		}

		LVCOLUMN insertionData = {0};
		insertionData.mask = LVCF_FMT | LVCF_TEXT | LVCF_SUBITEM | LVCF_NOTIFYSHELLLISTVIEW;
		insertionData.fmt = pDetails->columnFormat;
		if(pDetails->columnWidth >= 0) {
			insertionData.mask |= LVCF_WIDTH;
			insertionData.cx = pDetails->columnWidth;
		}
		insertionData.pszText = pDetails->pColumnText;
		insertionData.iSubItem = pDetails->columnSubItemIndex;
		int columnIndex = static_cast<int>(SendMessage(LVM_INSERTCOLUMN, (pDetails->insertAt == -1 ? 0x7FFFFFFF : pDetails->insertAt), reinterpret_cast<LPARAM>(&insertionData)));
		if(columnIndex != -1) {
			if(containedSysHeader32.GetExStyle() & WS_EX_RTLREADING) {
				HDITEM column = {0};
				column.mask = HDI_FORMAT;
				containedSysHeader32.SendMessage(HDM_GETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
				if(containedSysHeader32.GetExStyle() & WS_EX_RTLREADING) {
					column.fmt |= HDF_RTLREADING;
				}
				containedSysHeader32.SendMessage(HDM_SETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
			}
			if(pDetails->insertAt == 0 && IsComctl32Version610OrNewer()) {
				if((pDetails->columnFormat & LVCFMT_FIXED_WIDTH) == LVCFMT_FIXED_WIDTH || (pDetails->columnFormat & LVCFMT_SPLITBUTTON) == LVCFMT_SPLITBUTTON) {
					/* HACK: Inserting a column with the LVCFMT_FIXED_WIDTH or LVCFMT_SPLITBUTTON flag at position 0
									 doesn't work, so set this flag after insertion. */
					LVCOLUMN column = {0};
					column.mask = LVCF_FMT;
					SendMessage(LVM_GETCOLUMN, columnIndex, reinterpret_cast<LPARAM>(&column));
					column.fmt |= (pDetails->columnFormat & (LVCFMT_FIXED_WIDTH | LVCFMT_SPLITBUTTON));
					SendMessage(LVM_SETCOLUMN, columnIndex, reinterpret_cast<LPARAM>(&column));
				}
			}

			if(pDetails->columnWidth < 0) {
				// the column can't be auto-sized on creation, so do it _after_ insertion
				SendMessage(LVM_SETCOLUMNWIDTH, 0, MAKELPARAM(pDetails->columnWidth, 0));
			}
		}
		pDetails->insertedColumnID = ColumnIndexToID(columnIndex);
		return ((pDetails->insertedColumnID != -1) ? S_OK : E_FAIL);
	}

	HRESULT ExplorerListView::OnInternalColumnIDToIndex(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		PINT pColumnIndex = reinterpret_cast<PINT>(lParam);
		ATLASSERT_POINTER(pColumnIndex, INT);
		if(!pColumnIndex) {
			return E_POINTER;
		}

		*pColumnIndex = IDToColumnIndex(static_cast<LONG>(wParam));
		return (*pColumnIndex != -1 ? S_OK : E_FAIL);
	}

	HRESULT ExplorerListView::OnInternalColumnIndexToID(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		PLONG pColumnID = reinterpret_cast<PLONG>(lParam);
		ATLASSERT_POINTER(pColumnID, LONG);
		if(!pColumnID) {
			return E_POINTER;
		}

		*pColumnID = ColumnIndexToID(static_cast<int>(wParam));
		return (*pColumnID != -1 ? S_OK : E_FAIL);
	}

	HRESULT ExplorerListView::OnInternalGetColumnByID(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		int columnIndex = IDToColumnIndex(static_cast<LONG>(wParam));
		if(columnIndex != -1) {
			ClassFactory::InitColumn(columnIndex, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(lParam), FALSE);
			return S_OK;
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalSetSortArrow(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		CComPtr<IListViewColumn> pColumn = ClassFactory::InitColumn(static_cast<int>(wParam), this);
		if(pColumn) {
			return pColumn->put_SortArrow(static_cast<SortArrowConstants>(lParam));
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalGetColumnBitmap(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		CComPtr<IListViewColumn> pColumn = ClassFactory::InitColumn(static_cast<int>(wParam), this);
		if(pColumn) {
			return pColumn->get_BitmapHandle(reinterpret_cast<OLE_HANDLE*>(lParam));
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalSetColumnBitmap(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		CComPtr<IListViewColumn> pColumn = ClassFactory::InitColumn(static_cast<int>(wParam), this);
		if(pColumn) {
			return pColumn->put_BitmapHandle(static_cast<OLE_HANDLE>(lParam));
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalAddItem(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LPEXLVMADDITEMDATA pDetails = reinterpret_cast<LPEXLVMADDITEMDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXLVMADDITEMDATA);
		if(!pDetails) {
			return E_POINTER;
		}

		int itemIndex;
		if(wParam) {
			LVITEM insertionData = {0};
			insertionData.mask = LVIF_COLUMNS | LVIF_INDENT | LVIF_ISASHELLITEM | LVIF_PARAM | LVIF_STATE | LVIF_TEXT;
			insertionData.iItem = (pDetails->insertAt == -1 ? 0x7FFFFFFF : pDetails->insertAt);
			insertionData.pszText = LPSTR_TEXTCALLBACK;
			if(pDetails->iconIndex != -2) {
				insertionData.iImage = pDetails->iconIndex;
				insertionData.mask |= LVIF_IMAGE;
			}
			if(pDetails->overlayIndex > 0) {
				insertionData.state |= INDEXTOOVERLAYMASK(pDetails->overlayIndex);
				insertionData.stateMask |= LVIS_OVERLAYMASK;
			}
			insertionData.lParam = pDetails->itemData;
			insertionData.iIndent = pDetails->itemIndentation;
			if(pDetails->isGhosted != VARIANT_FALSE) {
				insertionData.state |= LVIS_CUT;
				insertionData.stateMask |= LVIS_CUT;
			}
			if(pDetails->groupID != -2) {
				insertionData.iGroupId = pDetails->groupID;
				insertionData.mask |= LVIF_GROUPID;
			}
			insertionData.cColumns = I_COLUMNSCALLBACK;
			itemIndex = static_cast<int>(SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
		} else {
			CComObject<ListViewItems>* pLvwItemsObj = NULL;
			CComObject<ListViewItems>::CreateInstance(&pLvwItemsObj);
			pLvwItemsObj->AddRef();
			pLvwItemsObj->SetOwner(this);
			itemIndex = pLvwItemsObj->Add(pDetails->pItemText, pDetails->insertAt, pDetails->iconIndex, pDetails->overlayIndex, pDetails->itemIndentation, pDetails->itemData, pDetails->isGhosted, pDetails->groupID, pDetails->pTileViewColumns, TRUE);
			pLvwItemsObj->Release();
		}
		pDetails->insertedItemID = ItemIndexToID(itemIndex);
		return ((pDetails->insertedItemID != -1) ? S_OK : E_FAIL);
	}

	HRESULT ExplorerListView::OnInternalItemIDToIndex(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		PINT pItemIndex = reinterpret_cast<PINT>(lParam);
		ATLASSERT_POINTER(pItemIndex, INT);
		if(!pItemIndex) {
			return E_POINTER;
		}

		*pItemIndex = IDToItemIndex(static_cast<LONG>(wParam));
		return (*pItemIndex != -1 ? S_OK : E_FAIL);
	}

	HRESULT ExplorerListView::OnInternalItemIndexToID(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		PLONG pItemID = reinterpret_cast<PLONG>(lParam);
		ATLASSERT_POINTER(pItemID, LONG);
		if(!pItemID) {
			return E_POINTER;
		}

		*pItemID = ItemIndexToID(static_cast<int>(wParam));
		return (*pItemID != -1 ? S_OK : E_FAIL);
	}

	HRESULT ExplorerListView::OnInternalGetItemByID(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		int itemIndex = IDToItemIndex(static_cast<LONG>(wParam));
		if(itemIndex != -1) {
			LVITEMINDEX i = {itemIndex, 0};
			ClassFactory::InitListItem(i, this, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(lParam), FALSE);
			return S_OK;
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalCreateItemContainer(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXLVMCREATEITEMCONTAINERDATA pDetails = reinterpret_cast<LPEXLVMCREATEITEMCONTAINERDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXLVMCREATEITEMCONTAINERDATA);
		if(!pDetails) {
			return E_POINTER;
		}
		ATLASSERT(pDetails->pItems);
		ATLASSERT(!pDetails->pContainer);

		CComObject<ListViewItemContainer>* pLvwItemContainerObj = NULL;
		CComObject<ListViewItemContainer>::CreateInstance(&pLvwItemContainerObj);
		pLvwItemContainerObj->AddRef();

		// clone all settings
		pLvwItemContainerObj->SetOwner(this);
		pLvwItemContainerObj->properties.useIndexes = ((GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA);
		// fast-add the items
		static_cast<IItemContainer*>(pLvwItemContainerObj)->AddItems(pDetails->numberOfItems, pDetails->pItems);

		HRESULT hr = pLvwItemContainerObj->QueryInterface(IID_PPV_ARGS(&pDetails->pContainer));
		RegisterItemContainer(static_cast<IItemContainer*>(pLvwItemContainerObj));
		pLvwItemContainerObj->Release();
		return hr;
	}

	HRESULT ExplorerListView::OnInternalGetItemIDsFromVariant(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXLVMGETITEMIDSDATA pDetails = reinterpret_cast<LPEXLVMGETITEMIDSDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXLVMGETITEMIDSDATA);
		if(!pDetails) {
			return E_POINTER;
		}
		ATLASSERT(pDetails->pVariant);
		ATLASSERT(!pDetails->pItemIDs);
		if(!pDetails->pVariant || pDetails->pItemIDs) {
			return E_INVALIDARG;
		}

		if(pDetails->pVariant->vt != VT_DISPATCH) {
			return DISP_E_TYPEMISMATCH;
		}

		// invalid arg - VB runtime error 380
		HRESULT hr = MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		if(pDetails->pVariant->pdispVal) {
			CComQIPtr<IListViewItem> pLvwItem = pDetails->pVariant->pdispVal;
			if(pLvwItem) {
				// a single ListViewItem object
				ATLASSUME(shellBrowserInterface.pInternalMessageListener);
				pDetails->pItemIDs = reinterpret_cast<PLONG>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_ALLOCATEMEMORY, sizeof(LONG), NULL));
				ATLASSERT_ARRAYPOINTER(pDetails->pItemIDs, LONG, 1);
				pDetails->pItemIDs[0] = -1;
				hr = pLvwItem->get_ID(&pDetails->pItemIDs[0]);
				pDetails->numberOfItems = 1;
			} else {
				CComQIPtr<IListViewItemContainer> pLvwItems = pDetails->pVariant->pdispVal;
				if(pLvwItems) {
					// a ListViewItemContainer collection
					CComQIPtr<IEnumVARIANT> pEnumerator = pLvwItems;
					LONG c = 0;
					pLvwItems->Count(&c);
					pDetails->numberOfItems = c;
					if(pDetails->numberOfItems > 0 && pEnumerator) {
						ATLASSUME(shellBrowserInterface.pInternalMessageListener);
						pDetails->pItemIDs = reinterpret_cast<PLONG>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_ALLOCATEMEMORY, pDetails->numberOfItems * sizeof(LONG), NULL));
						ATLASSERT_ARRAYPOINTER(pDetails->pItemIDs, LONG, pDetails->numberOfItems);
						VARIANT item;
						VariantInit(&item);
						ULONG dummy = 0;
						for(UINT i = 0; i < pDetails->numberOfItems && pEnumerator->Next(1, &item, &dummy) == S_OK; ++i) {
							pDetails->pItemIDs[i] = -1;
							if((item.vt == VT_DISPATCH) && item.pdispVal) {
								pLvwItem = item.pdispVal;
								if(pLvwItem) {
									pLvwItem->get_ID(&pDetails->pItemIDs[i]);
								}
							}
							VariantClear(&item);
						}
						hr = S_OK;
					}
				} else {
					CComQIPtr<IListViewItems> pLvwItems = pDetails->pVariant->pdispVal;
					if(pLvwItems) {
						// a ListViewItems collection
						CComQIPtr<IEnumVARIANT> pEnumerator = pLvwItems;
						LONG c = 0;
						pLvwItems->Count(&c);
						pDetails->numberOfItems = c;
						if(pDetails->numberOfItems > 0 && pEnumerator) {
							ATLASSUME(shellBrowserInterface.pInternalMessageListener);
							pDetails->pItemIDs = reinterpret_cast<PLONG>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_ALLOCATEMEMORY, pDetails->numberOfItems * sizeof(LONG), NULL));
							ATLASSERT_ARRAYPOINTER(pDetails->pItemIDs, LONG, pDetails->numberOfItems);
							VARIANT item;
							VariantInit(&item);
							ULONG dummy = 0;
							for(UINT i = 0; i < pDetails->numberOfItems && pEnumerator->Next(1, &item, &dummy) == S_OK; ++i) {
								pDetails->pItemIDs[i] = -1;
								if((item.vt == VT_DISPATCH) && item.pdispVal) {
									pLvwItem = item.pdispVal;
									if(pLvwItem) {
										pLvwItem->get_ID(&pDetails->pItemIDs[i]);
									}
								}
								VariantClear(&item);
							}
							hr = S_OK;
						}
					}
				}
			}
		}
		return hr;
	}

	HRESULT ExplorerListView::OnInternalSetItemIconIndex(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LVITEM item = {0};
		item.iItem = IDToItemIndex(static_cast<LONG>(wParam));
		item.iImage = static_cast<int>(lParam);
		item.mask = LVIF_IMAGE;
		/* NOTE: Don't use SendMessage here, because this leads to the message being processed by ShellListView
		 *       which won't know that item.iImage is an ID.
		 */
		if(DefWindowProc(LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return S_OK;
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalGetItemPosition(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		if(!lParam) {
			return E_POINTER;
		}

		LVITEMINDEX itemIndex = {IDToItemIndex(static_cast<LONG>(wParam)), 0};
		CComPtr<IListViewItem> pItem = ClassFactory::InitListItem(itemIndex, this);
		if(pItem) {
			LPPOINT pt = reinterpret_cast<LPPOINT>(lParam);
			return pItem->GetPosition(&pt->x, &pt->y);
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalSetItemPosition(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		if(!lParam) {
			return E_POINTER;
		}

		LVITEMINDEX itemIndex = {IDToItemIndex(static_cast<LONG>(wParam)), 0};
		CComPtr<IListViewItem> pItem = ClassFactory::InitListItem(itemIndex, this);
		if(pItem) {
			LPPOINT pt = reinterpret_cast<LPPOINT>(lParam);
			return pItem->SetPosition(pt->x, pt->y);
		}
		return DISP_E_BADINDEX;
	}

	HRESULT ExplorerListView::OnInternalMoveDraggedItems(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPSIZE pDistance = reinterpret_cast<LPSIZE>(lParam);
		if(!pDistance) {
			return E_POINTER;
		}
		if(!dragDropStatus.pDraggedItems) {
			return E_FAIL;
		}

		CComPtr<IUnknown> pEnum = NULL;
		HRESULT hr = dragDropStatus.pDraggedItems->get__NewEnum(&pEnum);
		if(SUCCEEDED(hr)) {
			ATLASSUME(pEnum);
			CComQIPtr<IEnumVARIANT> pItemEnumerator = pEnum;
			ATLASSUME(pItemEnumerator);
			VARIANT v;
			POINT pt;
			while(pItemEnumerator->Next(1, &v, NULL) == S_OK) {
				ATLASSERT(v.vt == VT_DISPATCH);
				CComQIPtr<IListViewItem> pItem = v.pdispVal;
				ATLASSUME(pItem);
				if(pItem) {
					hr = pItem->GetPosition(&pt.x, &pt.y);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						hr = pItem->SetPosition(pt.x + pDistance->cx, pt.y + pDistance->cy);
						ATLASSERT(SUCCEEDED(hr));
					}
				}
				VariantClear(&v);
			}
		}
		return hr;
	}

	HRESULT ExplorerListView::OnInternalHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXLVMHITTESTDATA pDetails = reinterpret_cast<LPEXLVMHITTESTDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXLVMHITTESTDATA);
		if(!pDetails) {
			return E_POINTER;
		}

		UINT hitTestDetails = 0;
		LVITEMINDEX itemIndex;
		int subItemIndex = -1;
		HitTest(pDetails->pointToTest.x, pDetails->pointToTest.y, &hitTestDetails, &itemIndex, &subItemIndex);

		pDetails->hitSubItemIndex = subItemIndex;
		pDetails->hitItemID = ItemIndexToID(itemIndex.iItem);
		pDetails->hitTestDetails = LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1);

		return S_OK;
	}

	HRESULT ExplorerListView::OnInternalSortItems(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/)
	{
		if(wParam != static_cast<WPARAM>(-1)) {
			return SortItems(sobShell, sobText, sobNone, sobNone, sobNone, _variant_t(static_cast<LONG>(wParam), VT_I4));
		} else {
			return SortItems();
		}
	}

	HRESULT ExplorerListView::OnInternalGetSortOrder(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return get_SortOrder(reinterpret_cast<SortOrderConstants*>(lParam));
	}

	HRESULT ExplorerListView::OnInternalSetSortOrder(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return put_SortOrder(static_cast<SortOrderConstants>(lParam));
	}

	HRESULT ExplorerListView::OnInternalGetClosestInsertMarkPosition(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LPPOINT pt = reinterpret_cast<LPPOINT>(wParam);
		InsertMarkPositionConstants* pPosition = &reinterpret_cast<InsertMarkPositionConstants*>(lParam)[1];
		CComPtr<IListViewItem> pItem = NULL;
		HRESULT hr = GetClosestInsertMarkPosition(pt->x, pt->y, pPosition, &pItem);
		if(pItem) {
			pItem->get_Index(&reinterpret_cast<PLONG>(lParam)[0]);
		} else {
			reinterpret_cast<PLONG>(lParam)[0] = -1;
		}
		return hr;
	}

	HRESULT ExplorerListView::OnInternalSetInsertMark(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LVITEMINDEX itemIndex = {static_cast<int>(wParam), 0};
		CComPtr<IListViewItem> pItem = ClassFactory::InitListItem(itemIndex, this);
		return SetInsertMarkPosition(static_cast<InsertMarkPositionConstants>(lParam), pItem);
	}

	HRESULT ExplorerListView::OnInternalSetDropHilitedItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return putref_DropHilitedItem(reinterpret_cast<IListViewItem*>(lParam));
	}

	HRESULT ExplorerListView::OnInternalControlIsDragSource(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		return dragDropStatus.IsDragging() ? S_OK : S_FALSE;
	}

	HRESULT ExplorerListView::OnInternalFireDragDropEvent(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		LPSHLVMDRAGDROPEVENTDATA pDetails = reinterpret_cast<LPSHLVMDRAGDROPEVENTDATA>(lParam);
		ATLASSERT(pDetails);
		if(!pDetails) {
			return E_POINTER;
		}
		if(pDetails->structSize < sizeof(SHLVMDRAGDROPEVENTDATA)) {
			return E_INVALIDARG;
		}

		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(pDetails->keyState, button, shift);
		POINT cursorPosition = {pDetails->cursorPosition.x, pDetails->cursorPosition.y};
		POINT cursorPositionHeader = cursorPosition;
		ScreenToClient(&cursorPosition);
		if(pDetails->headerIsTarget) {
			containedSysHeader32.ScreenToClient(&cursorPositionHeader);
		}
		if(pDetails->pDataObject) {
			switch(wParam) {
				case OLEDRAGEVENT_DRAGENTER:
					if(pDetails->headerIsTarget) {
						return Fire_HeaderOLEDragEnter(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<IListViewColumn**>(&pDetails->pDropTarget), button, shift, cursorPositionHeader.x, cursorPositionHeader.y, cursorPosition.x, cursorPosition.y, static_cast<HeaderHitTestConstants>(pDetails->hitTestDetails), &pDetails->autoHScrollVelocity, &pDetails->autoVScrollVelocity);
					} else {
						return Fire_OLEDragEnter(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<IListViewItem**>(&pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, static_cast<HitTestConstants>(pDetails->hitTestDetails), &pDetails->autoHScrollVelocity, &pDetails->autoVScrollVelocity);
					}
					break;
				case OLEDRAGEVENT_DRAGOVER:
					if(pDetails->headerIsTarget) {
						return Fire_HeaderOLEDragMouseMove(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<IListViewColumn**>(&pDetails->pDropTarget), button, shift, cursorPositionHeader.x, cursorPositionHeader.y, cursorPosition.x, cursorPosition.y, static_cast<HeaderHitTestConstants>(pDetails->hitTestDetails), &pDetails->autoHScrollVelocity, &pDetails->autoVScrollVelocity);
					} else {
						return Fire_OLEDragMouseMove(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<IListViewItem**>(&pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, static_cast<HitTestConstants>(pDetails->hitTestDetails), &pDetails->autoHScrollVelocity, &pDetails->autoVScrollVelocity);
					}
					break;
				case OLEDRAGEVENT_DROP:
					if(pDetails->headerIsTarget) {
						return Fire_HeaderOLEDragDrop(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<IListViewColumn*>(pDetails->pDropTarget), button, shift, cursorPositionHeader.x, cursorPositionHeader.y, cursorPosition.x, cursorPosition.y, static_cast<HeaderHitTestConstants>(pDetails->hitTestDetails));
					} else {
						return Fire_OLEDragDrop(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<OLEDropEffectConstants*>(&pDetails->effect), reinterpret_cast<IListViewItem**>(&pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, static_cast<HitTestConstants>(pDetails->hitTestDetails));
					}
					break;
				case OLEDRAGEVENT_DRAGLEAVE:
					if(pDetails->headerIsTarget) {
						return Fire_HeaderOLEDragLeave(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<IListViewColumn*>(pDetails->pDropTarget), button, shift, cursorPositionHeader.x, cursorPositionHeader.y, cursorPosition.x, cursorPosition.y, static_cast<HeaderHitTestConstants>(pDetails->hitTestDetails));
					} else {
						return Fire_OLEDragLeave(reinterpret_cast<IOLEDataObject*>(pDetails->pDataObject), reinterpret_cast<IListViewItem*>(pDetails->pDropTarget), button, shift, cursorPosition.x, cursorPosition.y, static_cast<HitTestConstants>(pDetails->hitTestDetails));
					}
					break;
			}
		}
		return E_INVALIDARG;
	}

	HRESULT ExplorerListView::OnInternalOLEDrag(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		LPEXLVMOLEDRAGDATA pDetails = reinterpret_cast<LPEXLVMOLEDRAGDATA>(lParam);
		ATLASSERT_POINTER(pDetails, EXLVMOLEDRAGDATA);
		if(!pDetails) {
			return E_POINTER;
		}
		ATLASSERT(pDetails->pDataObject);
		ATLASSERT(pDetails->pDraggedItems);

		if(pDetails->useSHDoDragDrop) {
			if(!properties.supportOLEDragImages) {
				// SHDoDragDrop always is with drag images
				return E_FAIL;
			}
			pDetails->performedEffects = odeNone;

			CComPtr<IOLEDataObject> pOLEDataObject = NULL;
			CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
			CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
			pOLEDataObjectObj->AddRef();
			pOLEDataObjectObj->Attach(pDetails->pDataObject);
			pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
			pOLEDataObjectObj->Release();
			pDetails->pDataObject->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));

			if(pDetails->pDraggedItems) {
				CComQIPtr<IListViewItemContainer> pContainer = pDetails->pDraggedItems;
				if(pContainer) {
					pContainer->Clone(&dragDropStatus.pDraggedItems);
				}
			}
			CComPtr<IDropSource> pDragSource = NULL;
			if(!pDetails->useShellDropSource) {
				QueryInterface(IID_PPV_ARGS(&pDragSource));
			}

			if(IsLeftMouseButtonDown()) {
				dragDropStatus.draggingMouseButton = MK_LBUTTON;
			} else if(IsRightMouseButtonDown()) {
				dragDropStatus.draggingMouseButton = MK_RBUTTON;
			}

			dragDropStatus.headerIsSource = FALSE;
			if(pOLEDataObject) {
				Raise_OLEStartDrag(pOLEDataObject);
			}
			HRESULT hr = SHDoDragDrop((RunTimeHelper::IsVista() ? NULL : static_cast<HWND>(*this)), pDetails->pDataObject, pDragSource, pDetails->supportedEffects, &pDetails->performedEffects);
			if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
				Raise_OLECompleteDrag(pOLEDataObject, static_cast<OLEDropEffectConstants>(pDetails->performedEffects));
			}

			dragDropStatus.Reset();
			return (hr != DRAGDROP_S_DROP && hr != DRAGDROP_S_CANCEL ? hr : S_OK);
		} else {
			pDetails->performedEffects = odeNone;
			CComQIPtr<IListViewItemContainer> pDraggedItems = pDetails->pDraggedItems;
			return OLEDrag(reinterpret_cast<LONG*>(&pDetails->pDataObject), static_cast<OLEDropEffectConstants>(pDetails->supportedEffects), static_cast<OLE_HANDLE>(-1), pDraggedItems, -1, reinterpret_cast<OLEDropEffectConstants*>(&pDetails->performedEffects));
		}
	}

	HRESULT ExplorerListView::OnInternalGetImageList(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		OLE_HANDLE h;
		HRESULT hr = get_hImageList(static_cast<ImageListConstants>(wParam), &h);
		*reinterpret_cast<HIMAGELIST*>(lParam) = reinterpret_cast<HIMAGELIST>(LongToHandle(h));
		return hr;
	}

	HRESULT ExplorerListView::OnInternalSetImageList(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		return put_hImageList(static_cast<ImageListConstants>(wParam), HandleToLong(reinterpret_cast<HIMAGELIST>(lParam)));
	}

	HRESULT ExplorerListView::OnInternalGetViewMode(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return get_View(reinterpret_cast<ViewConstants*>(lParam));
	}

	HRESULT ExplorerListView::OnInternalGetTileViewItemLines(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return get_TileViewItemLines(reinterpret_cast<LONG*>(lParam));
	}

	HRESULT ExplorerListView::OnInternalGetColumnHeaderVisibility(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam)
	{
		return get_ColumnHeaderVisibility(reinterpret_cast<ColumnHeaderVisibilityConstants*>(lParam));
	}

	HRESULT ExplorerListView::OnInternalIsItemVisible(UINT /*message*/, WPARAM wParam, LPARAM lParam)
	{
		if(!lParam) {
			return E_POINTER;
		}

		LVITEMINDEX itemIndex = {IDToItemIndex(static_cast<LONG>(wParam)), 0};
		CComPtr<IListViewItem> pItem = ClassFactory::InitListItem(itemIndex, this);
		if(pItem) {
			VARIANT_BOOL visible = VARIANT_FALSE;
			HRESULT hr = pItem->IsVisible(&visible);
			if(SUCCEEDED(hr)) {
				*reinterpret_cast<LPBOOL>(lParam) = VARIANTBOOL2BOOL(visible);
				return hr;
			}
		}
		return DISP_E_BADINDEX;
	}
#endif


LRESULT ExplorerListView::OnClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);

	if(!(properties.disabledEvents & deListClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_Click(button, shift, mousePosition.x, mousePosition.y);
	}

	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, NULL, NULL, TRUE, FALSE);
	// NOTE: The following line is NO typo.
	if(hitTestDetails != LVHT_ONITEMSTATEICON) {
		// SysListView32 seems to eat all WM_LBUTTONUP messages
		BOOL dummy = TRUE;
		OnLButtonUp(0, GetButtonStateBitField(), MAKELPARAM(mousePosition.x, mousePosition.y), dummy);
	}

	return 0;
}

LRESULT ExplorerListView::OnDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);
		Raise_DblClick(button, shift, mousePosition.x, mousePosition.y);
	}
	return 0;
}

LRESULT ExplorerListView::OnRClickNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(flags.ignoreRClickNotification) {
		return 0;
	}

	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);

	if(!(properties.disabledEvents & deListClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RClick(button, shift, mousePosition.x, mousePosition.y);
	}

	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, NULL, NULL, TRUE, FALSE);
	// NOTE: The following line is NO typo.
	if(hitTestDetails != LVHT_ONITEMSTATEICON) {
		// SysListView32 seems to eat all WM_RBUTTONUP messages
		BOOL dummy = TRUE;
		OnRButtonUp(0, GetButtonStateBitField(), MAKELPARAM(mousePosition.x, mousePosition.y), dummy);
	}

	return 0;
}

LRESULT ExplorerListView::OnRDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 2/*MouseButtonConstants.vbRightButton*/;
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);
		Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y);
	}
	return 0;
}

LRESULT ExplorerListView::OnSetFocusNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& wasHandled)
{
	if(!IsInDesignMode()) {
		// now that we've the focus, we should configure the IME
		IMEModeConstants ime = GetCurrentIMEContextMode(*this);
		if(ime != imeInherit) {
			ime = GetEffectiveIMEMode();
			if(ime != imeNoControl) {
				SetCurrentIMEContextMode(*this, ime);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ExplorerListView::OnAsyncDrawnNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLVASYNCDRAWN* pDetails = reinterpret_cast<NMLVASYNCDRAWN*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLVASYNCDRAWN);
	if(!pDetails) {
		return 0;
	}

	switch(pDetails->iPart) {
		case LVADPART_ITEM: {
			LVITEMINDEX itemIndex = {pDetails->iItem, 0};
			CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
			CComPtr<IListViewSubItem> pLvwSubItem = NULL;
			if(pDetails->iSubItem > 0) {
				pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, pDetails->iSubItem, this, FALSE, FALSE);
			}
			Raise_ItemAsynchronousDrawFailed(pLvwItem, pLvwSubItem, pDetails);
			return 0;
			break;
		}
		case LVADPART_GROUP: {
			CComPtr<IListViewGroup> pLvwGroup = ClassFactory::InitGroup(pDetails->iItem, this, FALSE);
			Raise_GroupAsynchronousDrawFailed(pLvwGroup, pDetails);
			return 0;
			break;
		}
	}
	return 0;
}

LRESULT ExplorerListView::OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return 0;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = pDetails->ptAction;
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, NULL, NULL, TRUE, TRUE, FALSE);

	LVITEMINDEX itemIndex = {pDetails->iItem, 0};
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pDetails->iSubItem > 0) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, pDetails->iSubItem, this, FALSE, FALSE);
	}
	Raise_ItemBeginDrag(pLvwItem, pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, pDetails->iSubItem != 0));

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_BEGINDRAG, FALSE, reinterpret_cast<LPARAM>(pNotificationDetails));
		}
	#endif

	return 0;
}

LRESULT ExplorerListView::OnBeginLabelEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	VARIANT_BOOL cancel = VARIANT_FALSE;

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			UINT message = (pNotificationDetails->code == LVN_BEGINLABELEDITA ? SHLVM_BEGINLABELEDITA : SHLVM_BEGINLABELEDITW);
			HRESULT hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(message, VARIANTBOOL2BOOL(cancel), reinterpret_cast<LPARAM>(pNotificationDetails));
			cancel = BOOL2VARIANTBOOL(hr != S_OK);
		}
		if(lstrlen(labelEditStatus.pTextToSetOnBegin) > 0) {
			containedEdit.SetWindowText(labelEditStatus.pTextToSetOnBegin);
			labelEditStatus.pTextToSetOnBegin[0] = TEXT('\0');
		}
	#endif

	LVITEMINDEX itemIndex = {reinterpret_cast<NMLVDISPINFO*>(pNotificationDetails)->item.iItem, 0};
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
	Raise_StartingLabelEditing(pLvwItem, &cancel);
	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnBeginRDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return 0;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = pDetails->ptAction;
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, NULL, NULL, TRUE, TRUE, FALSE);

	LVITEMINDEX itemIndex = {pDetails->iItem, 0};
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pDetails->iSubItem > 0) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, pDetails->iSubItem, this, FALSE, FALSE);
	}
	Raise_ItemBeginRDrag(pLvwItem, pLvwSubItem, button, shift, mousePosition.x, mousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, pDetails->iSubItem != 0));

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_BEGINDRAG, TRUE, reinterpret_cast<LPARAM>(pNotificationDetails));
		}
	#endif

	return 0;
}

LRESULT ExplorerListView::OnBeginScrollNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	Raise_BeforeScroll(reinterpret_cast<LPNMLVSCROLL>(pNotificationDetails)->dx, reinterpret_cast<LPNMLVSCROLL>(pNotificationDetails)->dy);
	return 0;
}

LRESULT ExplorerListView::OnColumnClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return 0;
	}

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iSubItem, this, FALSE);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 1/*MouseButtonConstants.vbLeftButton*/;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	containedSysHeader32.ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	Raise_ColumnClick(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	return 0;
}

LRESULT ExplorerListView::OnColumnDropDownNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return 0;
	}

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iSubItem, this, FALSE);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 1/*MouseButtonConstants.vbLeftButton*/;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	containedSysHeader32.ScreenToClient(&mousePosition);

	#ifdef INCLUDESHELLBROWSERINTERFACE
		POINT menuPosition = {-1, -1};
		RECT dropDownRectangle = {0};
		if(containedSysHeader32.SendMessage(HDM_GETITEMDROPDOWNRECT, pDetails->iSubItem, reinterpret_cast<LPARAM>(&dropDownRectangle))) {
			menuPosition.x = dropDownRectangle.left;
			menuPosition.y = dropDownRectangle.bottom + 1;
			containedSysHeader32.ClientToScreen(&menuPosition);
		}
	#endif
	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_ColumnDropDown(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener && menuPosition.x != -1 && menuPosition.y != -1) {
			if(showDefaultMenu != VARIANT_FALSE) {
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HEADERDROPDOWNMENU, ColumnIndexToID(pDetails->iSubItem), reinterpret_cast<LPARAM>(&menuPosition));
			}
		}
	#endif
	return 0;
}

LRESULT ExplorerListView::OnColumnOverflowClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iSubItem, this, FALSE);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button |= 1/*MouseButtonConstants.vbLeftButton*/;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	containedSysHeader32.ScreenToClient(&mousePosition);

	VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
	Raise_HeaderChevronClick(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, &showDefaultMenu);

	if(showDefaultMenu != VARIANT_FALSE) {
		RECT chevronRectangle = {0};
		if(containedSysHeader32.SendMessage(HDM_GETOVERFLOWRECT, 0, reinterpret_cast<LPARAM>(&chevronRectangle))) {
			POINT menuPosition;
			menuPosition.x = chevronRectangle.right;
			menuPosition.y = chevronRectangle.top;
			containedSysHeader32.ClientToScreen(&menuPosition);

			CMenu menu;
			if(menu.CreatePopupMenu()) {
				HDITEM column = {0};
				column.mask = HDI_TEXT;
				column.cchTextMax = MAX_COLUMNCAPTIONLENGTH;
				column.pszText = reinterpret_cast<LPTSTR>(malloc((column.cchTextMax + 1) * sizeof(TCHAR)));
				if(column.pszText) {
					CMenuItemInfo menuItemInfo;
					menuItemInfo.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
					menuItemInfo.fState = MFS_ENABLED;
					menuItemInfo.dwTypeData = column.pszText;
					int numberOfColumns = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0));
					for(int i = pDetails->iSubItem; i < numberOfColumns; ++i) {
						menuItemInfo.wID = i;
						if(containedSysHeader32.SendMessage(HDM_GETITEM, i, reinterpret_cast<LPARAM>(&column))) {
							menuItemInfo.cch = lstrlen(menuItemInfo.dwTypeData);
							menu.InsertMenuItem(i, TRUE, &menuItemInfo);
						}
					}
					free(column.pszText);
				}
				if(menu.GetMenuItemCount() > 0) {
					UINT chosenCommand = menu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, menuPosition.x, menuPosition.y, containedSysHeader32);
					if(chosenCommand != 0) {
						NMHEADER details = {0};
						details.hdr.code = HDN_ITEMCLICK;
						details.hdr.hwndFrom = containedSysHeader32;
						details.hdr.idFrom = containedSysHeader32.GetDlgCtrlID();
						details.iItem = chosenCommand;
						SendMessage(WM_NOTIFY, 0, reinterpret_cast<LPARAM>(&details));
					}
				}
			}
		}
	}
	return 0;
}

LRESULT ExplorerListView::OnDeleteAllItemsNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(properties.disabledEvents & deFreeItemData) {
		// no need for wasting time
		if(!RunTimeHelper::IsCommCtrl6()) {
			#ifdef USE_STL
				itemIDs.clear();
				itemParams.clear();
			#else
				itemIDs.RemoveAll();
				itemParams.RemoveAll();
			#endif
		}
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pInternalMessageListener) {
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_FREEITEM, static_cast<WPARAM>(-1), 0);
			}
		#endif
		return TRUE;
	} else {
		// we want a LVN_DELETEITEM notification for each item
		return FALSE;
	}
}

LRESULT ExplorerListView::OnDeleteItemNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return 0;
	}

	if(!(properties.disabledEvents & deFreeItemData)) {
		LVITEMINDEX itemIndex = {pDetails->iItem, 0};
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
		Raise_FreeItemData(pLvwItem);
	}

	BOOL virtualMode = ((GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA);
	RemoveItemFromItemContainers(virtualMode ? pDetails->iItem : flags.itemIDBeingRemoved);
	if(!RunTimeHelper::IsCommCtrl6() && !virtualMode) {
		#ifdef USE_STL
			itemIDs.erase(itemIDs.begin() + pDetails->iItem);
			std::list<ItemData>::iterator iter2 = itemParams.begin();
			if(iter2 != itemParams.end()) {
				std::advance(iter2, pDetails->iItem);
				if(iter2 != itemParams.end()) {
					itemParams.erase(iter2);
				}
			}
		#else
			itemIDs.RemoveAt(pDetails->iItem);
			POSITION p = itemParams.FindIndex(pDetails->iItem);
			if(p) {
				itemParams.RemoveAt(p);
			}
		#endif
	}
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_FREEITEM, flags.itemIDBeingRemoved, 0);
		}
	#endif
	flags.itemIDBeingRemoved = -1;

	return 0;
}

LRESULT ExplorerListView::OnEndLabelEditNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLVDISPINFO* pDetails = reinterpret_cast<NMLVDISPINFO*>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_ENDLABELEDITA) {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOW);
		}
	#endif

	if(pDetails->item.pszText) {
		LVITEMINDEX itemIndex = {pDetails->item.iItem, pDetails->item.iGroup};
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComBSTR newItemText;
		if(pNotificationDetails->code == LVN_ENDLABELEDITA) {
			newItemText = reinterpret_cast<LPNMLVDISPINFOA>(pDetails)->item.pszText;
		} else {
			newItemText = reinterpret_cast<LPNMLVDISPINFOW>(pDetails)->item.pszText;
		}
		BSTR itemText = NULL;
		pLvwItem->get_Text(&itemText);

		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_RenamingItem(pLvwItem, itemText, newItemText, &cancel);
		if(cancel == VARIANT_FALSE) {
			labelEditStatus.editedItem = itemIndex;
			labelEditStatus.previousText = itemText;
			labelEditStatus.acceptText = TRUE;

			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(shellBrowserInterface.pInternalMessageListener) {
					WCHAR pNewName[MAX_ITEMTEXTLENGTH + 1];
					lstrcpynW(pNewName, newItemText, MAX_ITEMTEXTLENGTH + 1);
					#ifdef UNICODE
						lstrcpyn(labelEditStatus.pTextToSetOnBegin, pNewName, MAX_ITEMTEXTLENGTH + 1);
					#else
						CW2A converter(pNewName);
						lstrcpyn(labelEditStatus.pTextToSetOnBegin, converter, MAX_ITEMTEXTLENGTH + 1);
					#endif
					HRESULT hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_RENAMEDITEM, ItemIndexToID(labelEditStatus.editedItem.iItem), reinterpret_cast<LPARAM>(pNewName));
					labelEditStatus.acceptText = (hr == S_OK);
					if(labelEditStatus.acceptText) {
						labelEditStatus.pTextToSetOnBegin[0] = TEXT('\0');
					}
				}
			#endif
			SysFreeString(itemText);
			return TRUE;
		}
		SysFreeString(itemText);
	}
	return FALSE;
}

LRESULT ExplorerListView::OnEndScrollNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	Raise_AfterScroll(reinterpret_cast<LPNMLVSCROLL>(pNotificationDetails)->dx, reinterpret_cast<LPNMLVSCROLL>(pNotificationDetails)->dy);
	return 0;
}

LRESULT ExplorerListView::OnGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(properties.disabledEvents & deItemGetDisplayInfo) {
		return 0;
	}

	NMLVDISPINFO* pDetails = reinterpret_cast<NMLVDISPINFO*>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_GETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOW);
		}
	#endif

	LVITEMINDEX itemIndex = {pDetails->item.iItem, 0};
	if(IsComctl32Version610OrNewer()) {
		// TODO: On Vista, pDetails->item.iGroup seems to be always 0.
		ATLASSERT(pDetails->item.iGroup == 0);
		itemIndex.iGroup = (pDetails->item.iGroup > 0 ? pDetails->item.iGroup : 0);
	}
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pLvwItem && pDetails->item.iSubItem > 0) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, pDetails->item.iSubItem, this, FALSE, FALSE);
	}
	LONG iconIndex = -2;
	LONG indent = 0;
	LONG groupID = I_GROUPIDNONE;
	CComBSTR itemText = L"";
	LONG maxItemTextLength = 0;
	LONG overlayIndex = 0;
	LONG stateImageIndex = 0;
	ItemStateConstants itemStates = static_cast<ItemStateConstants>(0);

	LPSAFEARRAY pTileViewColumns = NULL;
	if((pDetails->item.mask & LVIF_COLUMNS) == LVIF_COLUMNS || (pDetails->item.mask & LVIF_COLFMT) == LVIF_COLFMT) {
		if(pDetails->item.cColumns >= 0 && pDetails->item.cColumns <= 50) {
			CComPtr<IRecordInfo> pRecordInfo = NULL;
			CLSID clsidTILEVIEWSUBITEM = {0};
			#ifdef UNICODE
				CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
			#else
				CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
				ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_ExLVwLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidTILEVIEWSUBITEM), &pRecordInfo)));
			#endif
			pTileViewColumns = SafeArrayCreateVectorEx(VT_RECORD, 0, pDetails->item.cColumns, pRecordInfo);
			ATLASSERT_POINTER(pTileViewColumns, SAFEARRAY);
		}
	}

	RequestedInfoConstants requestedInfo = static_cast<RequestedInfoConstants>(0);
	if(pDetails->item.mask & LVIF_IMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIconIndex);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				iconIndex = pDetails->item.iImage;
			}
		#endif
	}
	if(pDetails->item.mask & LVIF_INDENT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIndent);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				indent = pDetails->item.iIndent;
			}
		#endif
	}
	if(pDetails->item.mask & LVIF_GROUPID) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riGroupID);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				groupID = pDetails->item.iGroupId;
			}
		#endif
	}
	if((pDetails->item.mask & LVIF_COLUMNS) == LVIF_COLUMNS || (pDetails->item.mask & LVIF_COLFMT) == LVIF_COLFMT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riTileViewColumns);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(pDetails->item.cColumns > 0 && pTileViewColumns) {
				if(shellBrowserInterface.pMessageListener) {
					TILEVIEWSUBITEM element = {0};
					for(LONG i = 0; i < static_cast<LONG>(pDetails->item.cColumns); ++i) {
						element.ColumnIndex = pDetails->item.puColumns[i];
						element.BeginNewColumn = BOOL2VARIANTBOOL(pDetails->item.piColFmt[i] & LVCFMT_LINE_BREAK);
						element.FillRemainder = BOOL2VARIANTBOOL(pDetails->item.piColFmt[i] & LVCFMT_FILL);
						element.WrapToMultipleLines = BOOL2VARIANTBOOL(pDetails->item.piColFmt[i] & LVCFMT_WRAP);
						element.NoTitle = BOOL2VARIANTBOOL(pDetails->item.piColFmt[i] & LVCFMT_NO_TITLE);
						ATLVERIFY(SUCCEEDED(SafeArrayPutElement(pTileViewColumns, &i, &element)));
					}
				}
			}
		#endif
	}
	if(pDetails->item.mask & LVIF_TEXT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riItemText);
		// exclude the terminating 0
		maxItemTextLength = pDetails->item.cchTextMax - 1;
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				if(pNotificationDetails->code == LVN_GETDISPINFOA) {
					itemText = reinterpret_cast<LPNMLVDISPINFOA>(pDetails)->item.pszText;
				} else {
					itemText = reinterpret_cast<LPNMLVDISPINFOW>(pDetails)->item.pszText;
				}
			}
		#endif
	}
	if(pDetails->item.mask & LVIF_STATE) {
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(shellBrowserInterface.pMessageListener) {
				itemStates = static_cast<ItemStateConstants>(pDetails->item.state);
			}
		#endif
		if(pDetails->item.stateMask & LVIS_OVERLAYMASK) {
			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(shellBrowserInterface.pMessageListener) {
					overlayIndex = (pDetails->item.state & LVIS_OVERLAYMASK) >> 8;
				}
			#endif
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riOverlayIndex);
		}
		if(pDetails->item.stateMask & LVIS_STATEIMAGEMASK) {
			#ifdef INCLUDESHELLBROWSERINTERFACE
				if(shellBrowserInterface.pMessageListener) {
					stateImageIndex = (pDetails->item.state & LVIS_STATEIMAGEMASK) >> 12;
				}
			#endif
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riStateImageIndex);
		}
		if(pDetails->item.stateMask & LVIS_ACTIVATING) {
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riActivating);
		}
		if(pDetails->item.stateMask & LVIS_FOCUSED) {
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riCaret);
		}
		if(pDetails->item.stateMask & LVIS_DROPHILITED) {
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riDropHilited);
		}
		if(pDetails->item.stateMask & LVIS_CUT) {
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riGhosted);
		}
		if(pDetails->item.stateMask & LVIS_GLOW) {
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riGlowing);
		}
		if(pDetails->item.stateMask & LVIS_SELECTED) {
			requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riSelected);
		}
		#ifdef INCLUDESHELLBROWSERINTERFACE
			// TODO...
		#endif
	}

	VARIANT_BOOL dontAskAgain = VARIANT_FALSE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pMessageListener) {
			dontAskAgain = BOOL2VARIANTBOOL((pDetails->item.mask & LVIF_DI_SETITEM) == LVIF_DI_SETITEM);
		}
	#endif
	Raise_ItemGetDisplayInfo(pLvwItem, pLvwSubItem, requestedInfo, &iconIndex, &indent, &groupID, &pTileViewColumns, maxItemTextLength, &itemText, &overlayIndex, &stateImageIndex, &itemStates, &dontAskAgain);

	if(pDetails->item.mask & LVIF_IMAGE) {
		pDetails->item.iImage = iconIndex;
	}
	if(pDetails->item.mask & LVIF_INDENT) {
		pDetails->item.iIndent = indent;
	}
	if(pDetails->item.mask & LVIF_GROUPID) {
		pDetails->item.iGroupId = groupID;
	}
	if((pDetails->item.mask & LVIF_COLUMNS) == LVIF_COLUMNS || (pDetails->item.mask & LVIF_COLFMT) == LVIF_COLFMT) {
		if(pTileViewColumns) {
			LONG l;
			ATLVERIFY(SUCCEEDED(SafeArrayGetLBound(pTileViewColumns, 1, &l)));
			LONG u;
			ATLVERIFY(SUCCEEDED(SafeArrayGetUBound(pTileViewColumns, 1, &u)));

			TILEVIEWSUBITEM element = {0};
			for(LONG i = l; (i <= u) && (i - l < static_cast<LONG>(pDetails->item.cColumns)); ++i) {
				ATLVERIFY(SUCCEEDED(SafeArrayGetElement(pTileViewColumns, &i, &element)));
				pDetails->item.puColumns[i - l] = element.ColumnIndex;
				pDetails->item.piColFmt[i - l] = 0;
				if(element.BeginNewColumn != VARIANT_FALSE) {
					pDetails->item.piColFmt[i - l] |= LVCFMT_LINE_BREAK;
				}
				if(element.FillRemainder != VARIANT_FALSE) {
					pDetails->item.piColFmt[i - l] |= LVCFMT_FILL;
				}
				if(element.WrapToMultipleLines != VARIANT_FALSE) {
					pDetails->item.piColFmt[i - l] |= LVCFMT_WRAP;
				}
				if(element.NoTitle != VARIANT_FALSE) {
					pDetails->item.piColFmt[i - l] |= LVCFMT_NO_TITLE;
				}
			}
			pDetails->item.cColumns = max(u - l + 1, 0);
		}
	}

	if(pDetails->item.mask & LVIF_TEXT) {
		// don't forget the terminating 0
		int bufferSize = itemText.Length() + 1;
		if(bufferSize > pDetails->item.cchTextMax) {
			bufferSize = pDetails->item.cchTextMax;
		}
		if(pNotificationDetails->code == LVN_GETDISPINFOA) {
			if(bufferSize > 1) {
				lstrcpynA(reinterpret_cast<LPNMLVDISPINFOA>(pDetails)->item.pszText, CW2A(itemText), bufferSize);
			} else {
				reinterpret_cast<LPNMLVDISPINFOA>(pDetails)->item.pszText[0] = '\0';
			}
		} else {
			if(bufferSize > 1) {
				lstrcpynW(reinterpret_cast<LPNMLVDISPINFOW>(pDetails)->item.pszText, itemText, bufferSize);
			} else {
				reinterpret_cast<LPNMLVDISPINFOW>(pDetails)->item.pszText[0] = L'\0';
			}
		}
	}
	if(pDetails->item.mask & LVIF_STATE) {
		pDetails->item.state = itemStates;
		if(pDetails->item.stateMask & LVIS_OVERLAYMASK) {
			pDetails->item.state &= ~LVIS_OVERLAYMASK;
			pDetails->item.state |= INDEXTOOVERLAYMASK(overlayIndex);
		}
		if(pDetails->item.stateMask & LVIS_STATEIMAGEMASK) {
			pDetails->item.state &= ~LVIS_STATEIMAGEMASK;
			pDetails->item.state |= INDEXTOSTATEIMAGEMASK(stateImageIndex);
		}
	}

	if(dontAskAgain == VARIANT_FALSE) {
		pDetails->item.mask &= ~LVIF_DI_SETITEM;
	} else {
		pDetails->item.mask |= LVIF_DI_SETITEM;
	}

	if(pTileViewColumns) {
		SafeArrayDestroy(pTileViewColumns);
	}
	return 0;
}

LRESULT ExplorerListView::OnGetEmptyMarkupNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLVEMPTYMARKUP* pDetails = reinterpret_cast<NMLVEMPTYMARKUP*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLVEMPTYMARKUP);
	if(!pDetails) {
		return FALSE;
	}

	if(properties.pEmptyMarkupText) {
		ATLVERIFY(SUCCEEDED(StringCchCopyW(pDetails->szMarkup, L_MAX_URL_LENGTH - 1, properties.pEmptyMarkupText)));
		switch(properties.emptyMarkupTextAlignment) {
			case alCenter:
				pDetails->dwFlags = EMF_CENTERED;
				break;
		}
		return TRUE;
	}

	return FALSE;
}

LRESULT ExplorerListView::OnGetEmptyTextNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLVDISPINFO* pDetails = reinterpret_cast<NMLVDISPINFO*>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_GETEMPTYTEXTA) {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOW);
		}
	#endif

	if(pDetails->item.mask & LVIF_TEXT) {
		if(properties.pEmptyMarkupText) {
			if(pNotificationDetails->code == LVN_GETEMPTYTEXTA) {
				ATLVERIFY(SUCCEEDED(StringCchCopyA(reinterpret_cast<LPNMLVDISPINFOA>(pDetails)->item.pszText, pDetails->item.cchTextMax - 1, CW2A(properties.pEmptyMarkupText))));
			} else {
				ATLVERIFY(SUCCEEDED(StringCchCopyW(reinterpret_cast<LPNMLVDISPINFOW>(pDetails)->item.pszText, pDetails->item.cchTextMax - 1, properties.pEmptyMarkupText)));
			}
			return TRUE;
		}
	}

	return FALSE;
}

LRESULT ExplorerListView::OnGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_GETINFOTIPA) {
			ATLASSERT_POINTER(pDetails, NMLVGETINFOTIPA);
		} else {
			ATLASSERT_POINTER(pDetails, NMLVGETINFOTIPW);
		}
	#endif

	LVITEMINDEX itemIndex = {pDetails->iItem, 0};
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
	CComBSTR text = L"";
	DWORD style = static_cast<DWORD>(SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0));
	if(style & LVS_EX_LABELTIP) {
		// small hack to provide some more functionality (LVS_EX_INFOTIP seems to imply LVS_EX_LABELTIP)
		if(pNotificationDetails->code == LVN_GETINFOTIPA) {
			text.Append(reinterpret_cast<LPNMLVGETINFOTIPA>(pDetails)->pszText);
		} else {
			text.Append(reinterpret_cast<LPNMLVGETINFOTIPW>(pDetails)->pszText);
		}
	}
	// exclude the terminating 0
	LONG maxTextLength = pDetails->cchTextMax - 1;

	VARIANT_BOOL abortToolTip = VARIANT_FALSE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			if(shellBrowserInterface.shellBrowserBuildNumber > 440) {
				// send SHLVM_GETINFOTIPEX
				SHLVMGETINFOTIPDATA infoTipData = {0};
				infoTipData.structSize = sizeof(SHLVMGETINFOTIPDATA);
				infoTipData.prependText = TRUE;
				infoTipData.pPrependedText = OLE2W(text);
				infoTipData.prependedTextSize = lstrlenW(infoTipData.pPrependedText);
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_GETINFOTIPEX, static_cast<WPARAM>(ItemIndexToID(pDetails->iItem)), reinterpret_cast<LPARAM>(&infoTipData));

				if(!infoTipData.cancelToolTip) {
					if(lstrlenW(infoTipData.pShellInfoTipText) > 0) {
						if(text.Length() > 0) {
							if(infoTipData.prependText) {
								text.Append(TEXT("\r\n"));
							} else {
								text.Empty();
							}
						}

						text.Append(infoTipData.pShellInfoTipText, infoTipData.shellInfoTipTextSize);
					}
					CoTaskMemFree(infoTipData.pShellInfoTipText);
				} else {
					abortToolTip = VARIANT_TRUE;
				}
			} else {
				// send SHLVM_GETINFOTIP
				LPWSTR pInfoTip = NULL;
				shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_GETINFOTIP, static_cast<WPARAM>(ItemIndexToID(pDetails->iItem)), reinterpret_cast<LPARAM>(&pInfoTip));
				if(pInfoTip) {
					if(lstrlenW(pInfoTip) > 0) {
						if(text.Length() > 0) {
							text.Append(TEXT("\r\n"));
						}
						text.Append(pInfoTip);
					}
					CoTaskMemFree(pInfoTip);
				}
			}
		}
	#endif
	if(abortToolTip == VARIANT_FALSE) {
		Raise_ItemGetInfoTipText(pLvwItem, maxTextLength, &text, &abortToolTip);
	}
	flags.cancelToolTip = VARIANTBOOL2BOOL(abortToolTip);

	// don't forget the terminating 0
	int bufferSize = text.Length() + 1;
	if(bufferSize > pDetails->cchTextMax) {
		bufferSize = pDetails->cchTextMax;
	}
	if(pNotificationDetails->code == LVN_GETINFOTIPA) {
		int requiredSize = WideCharToMultiByte(CP_ACP, 0, text, bufferSize, "", 0, "", NULL);
		if(requiredSize > 0) {
			LPSTR pBuffer = reinterpret_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, (requiredSize + 1) * sizeof(CHAR)));
			if(pBuffer) {
				WideCharToMultiByte(CP_ACP, 0, text, bufferSize, pBuffer, requiredSize, "", NULL);
				/* On Vista and newer, many info tips contain Unicode characters that cannot be converted.
				 * We have told WideCharToMultiByte to replace them with \0. Now remove those \0.
				 */
				LPSTR pDst = reinterpret_cast<LPNMLVGETINFOTIPA>(pDetails)->pszText;
				ZeroMemory(pDst, pDetails->cchTextMax);
				int j = 0;
				for(int i = 0; i < min(bufferSize, requiredSize); i++) {
					if(pBuffer[i] != '\0') {
						pDst[j++] = pBuffer[i];
					}
				}
				HeapFree(GetProcessHeap(), 0, pBuffer);
			}
		}
	} else {
		lstrcpynW(reinterpret_cast<LPNMLVGETINFOTIPW>(pDetails)->pszText, text, bufferSize);
	}

	return 0;
}

LRESULT ExplorerListView::OnGroupInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGROUP pDetails = reinterpret_cast<LPNMLVGROUP>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLVGROUP);
	if(!pDetails) {
		return 0;
	}

	CComPtr<IListViewGroup> pGroup;
	if((pDetails->uNewState & LVGS_COLLAPSED) != (pDetails->uOldState & LVGS_COLLAPSED)) {
		ATLASSERT(pDetails->iGroupId != 0);
		if(!pGroup) {
			pGroup = ClassFactory::InitGroup(pDetails->iGroupId, this);
		}
		if(pDetails->uNewState & LVGS_COLLAPSED) {
			Raise_CollapsedGroup(pGroup);
		} else {
			Raise_ExpandedGroup(pGroup);
		}
	}
	if((pDetails->uNewState & LVGS_FOCUSED) != (pDetails->uOldState & LVGS_FOCUSED)) {
		ATLASSERT(pDetails->iGroupId != 0);
		if(!pGroup) {
			pGroup = ClassFactory::InitGroup(pDetails->iGroupId, this);
		}
		if(pDetails->uNewState & LVGS_FOCUSED) {
			Raise_GroupGotFocus(pGroup);
		} else {
			Raise_GroupLostFocus(pGroup);
		}
	}
	if((pDetails->uNewState & LVGS_SELECTED) != (pDetails->uOldState & LVGS_SELECTED)) {
		ATLASSERT(pDetails->iGroupId != 0);
		if(!pGroup) {
			pGroup = ClassFactory::InitGroup(pDetails->iGroupId, this);
		}
		Raise_GroupSelectionChanged(pGroup);
	}

	/* NOTE: LVGS_SUBSETED doesn't specify whether the group is actually subseted. It only marks the group
	         as being subsetable.
	         LVGS_SUBSETLINKFOCUSED doesn't seem to work. */

	return 0;
}

LRESULT ExplorerListView::OnHotTrackNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return FALSE;
	}

	// TODO: fully support LVITEMINDEX
	if(pDetails->iItem != cachedSettings.hotItem.iItem) {
		LVITEMINDEX newHotItem = {pDetails->iItem, 0};
		if(!(properties.disabledEvents & deHotItemChangeEvents)) {
			CComPtr<IListViewItem> pPreviousHotItem = ClassFactory::InitListItem(cachedSettings.hotItem, this);
			CComPtr<IListViewItem> pNewHotItem = ClassFactory::InitListItem(newHotItem, this);
			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_HotItemChanging(pPreviousHotItem, pNewHotItem, &cancel);
			if(cancel != VARIANT_FALSE) {
				return TRUE;
			}
		}
		cachedSettings.hotItemTmp = newHotItem;
	}

	return FALSE;
}

LRESULT ExplorerListView::OnIncrementalSearchNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListKeyboardEvents)) {
		LPNMLVFINDITEM pDetails = reinterpret_cast<LPNMLVFINDITEM>(pNotificationDetails);
		#ifdef DEBUG
			if(pNotificationDetails->code == LVN_INCREMENTALSEARCHA) {
				ATLASSERT_POINTER(pDetails, NMLVFINDITEMA);
			} else {
				ATLASSERT_POINTER(pDetails, NMLVFINDITEMW);
			}
		#endif

		LONG itemToSelect = -1;
		if(pNotificationDetails->code == LVN_INCREMENTALSEARCHA) {
			Raise_IncrementalSearching(_bstr_t(reinterpret_cast<LPNMLVFINDITEMA>(pDetails)->lvfi.psz).Detach(), &itemToSelect);
		} else {
			Raise_IncrementalSearching(_bstr_t(reinterpret_cast<LPNMLVFINDITEMW>(pDetails)->lvfi.psz).Detach(), &itemToSelect);
		}
		pDetails->lvfi.lParam = itemToSelect;
	}
	return 0;
}

LRESULT ExplorerListView::OnItemActivateNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMITEMACTIVATE pDetails = reinterpret_cast<LPNMITEMACTIVATE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMITEMACTIVATE);
	if(!pDetails) {
		return 0;
	}

	LVITEMINDEX itemIndex = {pDetails->iItem, 0};
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pDetails->iSubItem > 0) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, pDetails->iSubItem, this, FALSE, FALSE);
	}
	SHORT shift = 0;
	if(pDetails->uKeyFlags & LVKF_ALT) {
		shift |= 4 /*ShiftConstants.vbAltMask*/;
	}
	if(pDetails->uKeyFlags & LVKF_CONTROL) {
		shift |= 2 /*ShiftConstants.vbCtrlMask*/;
	}
	if(pDetails->uKeyFlags & LVKF_SHIFT) {
		shift |= 1 /*ShiftConstants.vbShiftMask*/;
	}
	// all other members of NMITEMACTIVATE are garbage

	Raise_ItemActivate(pLvwItem, pLvwSubItem, shift, pDetails->ptAction.x, pDetails->ptAction.y);

	return 0;
}

LRESULT ExplorerListView::OnItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLISTVIEW);
	if(!pDetails) {
		return 0;
	}

	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uOldState & LVIS_FOCUSED) != (pDetails->uNewState & LVIS_FOCUSED)) {
			ATLASSERT(pDetails->iItem != -1);
			caretChangedStatus.previousCaretItem = caretChangedStatus.newCaretItem;
			caretChangedStatus.newCaretItem.iItem = pDetails->iItem;
			if(caretChangedStatus.previousCaretItem.iItem != -1) {
				KillTimer(timers.ID_CARETCHANGED);
				Raise_CaretChanged(caretChangedStatus.previousCaretItem, caretChangedStatus.newCaretItem);
			} else {
				/* This was the first notification during this caret change. We wait some milliseconds for a
				   second one. If no second one is sent, either the previous or the new caret item is -1. */
				SetTimer(timers.ID_CARETCHANGED, timers.INT_CARETCHANGED);
			}
		}
		if((pDetails->uOldState & LVIS_SELECTED) != (pDetails->uNewState & LVIS_SELECTED)) {
			// pDetails->iItem will be -1 in virtual mode
			if(pDetails->iItem != -1) {
				LVITEMINDEX itemIndex = {pDetails->iItem, 0};
				Raise_ItemSelectionChanged(itemIndex);
			}
		}
	}

	return 0;
}

LRESULT ExplorerListView::OnKeyDownNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deListKeyboardEvents)) {
		BOOL isChar = FALSE;
		LPNMLVKEYDOWN pDetails = reinterpret_cast<LPNMLVKEYDOWN>(pNotificationDetails);
		ATLASSERT_POINTER(pDetails, NMLVKEYDOWN);
		if(!pDetails) {
			return FALSE;
		}

		////////////////////////////////////////////////////////////////
		// taken from wine\dlls\user32\message.c -> TranslateMessage() (Wine 0.9.52)
		WCHAR wp[2];
		BYTE state[256];

		GetKeyboardState(state);
		/* FIXME : should handle ToUnicode yielding 2 */
		switch(ToUnicode(pDetails->wVKey, HIWORD(lastKeyDownLParam), state, wp, 2, 0)) {
			case -1:
			case 1:
				isChar = TRUE;
				break;
		}
		////////////////////////////////////////////////////////////////

		VARIANT_BOOL cancel = VARIANT_FALSE;
		if(isChar) {
			BSTR searchString = SysAllocString(L"");
			get_IncrementalSearchString(&searchString);
			Raise_IncrementalSearchStringChanging(searchString, static_cast<SHORT>(pDetails->wVKey), &cancel);
			SysFreeString(searchString);
		}
		if(cancel != VARIANT_FALSE) {
			// exclude the char
			return TRUE;
		}
	}
	return FALSE;
}

LRESULT ExplorerListView::OnLinkClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	PNMLVLINK pDetails = reinterpret_cast<PNMLVLINK>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLVLINK);
	if(!pDetails) {
		return 0;
	}

	ATLASSERT(pDetails->iItem == -1);
	ATLASSERT(pDetails->link.mask == 0);
	ATLASSERT(pDetails->link.state == 0);
	ATLASSERT(pDetails->link.stateMask == 0);
	ATLASSERT(pDetails->link.szUrl[0] == L'\0');

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button = mouseStatus.lastMouseUpButton;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, NULL, NULL, TRUE, TRUE, FALSE);

	if(pDetails->iSubItem == -1 && _wcsicmp(pDetails->link.szID, L"EmptyMarkup") == 0) {
		Raise_EmptyMarkupTextLinkClick(pDetails->link.iLink, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
	} else {
		CComPtr<IListViewGroup> pLvwGroup = ClassFactory::InitGroup(pDetails->iSubItem, this, FALSE);
		Raise_GroupTaskLinkClick(pLvwGroup, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return 0;
}

LRESULT ExplorerListView::OnMarqueeBeginNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_BeginMarqueeSelection(&cancel);
	return (cancel != VARIANT_FALSE);
}

LRESULT ExplorerListView::OnODCacheHintNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVCACHEHINT pDetails = reinterpret_cast<LPNMLVCACHEHINT>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLVCACHEHINT);
	if(!pDetails) {
		return 0;
	}

	LVITEMINDEX firstItem = {pDetails->iFrom, 0};
	LVITEMINDEX lastItem = {pDetails->iTo, 0};
	CComPtr<IListViewItem> pFirstItem = ClassFactory::InitListItem(firstItem, this, FALSE);
	CComPtr<IListViewItem> pLastItem = ClassFactory::InitListItem(lastItem, this, FALSE);
	Raise_CacheItemsHint(pFirstItem, pLastItem);
	return 0;
}

LRESULT ExplorerListView::OnODFindItemNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVFINDITEM pDetails = reinterpret_cast<LPNMLVFINDITEM>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_ODFINDITEMA) {
			ATLASSERT_POINTER(pDetails, NMLVFINDITEMA);
		} else {
			ATLASSERT_POINTER(pDetails, NMLVFINDITEMW);
		}
	#endif

	LVITEMINDEX itemIndex = {pDetails->iStart, 0};
	CComPtr<IListViewItem> pItemToStartWith = ClassFactory::InitListItem(itemIndex, this);
	if(!pItemToStartWith) {
		// pDetails->iStart is probably too big
		if(pDetails->lvfi.lParam & LVFI_WRAP) {
			itemIndex.iItem = 0;
			pItemToStartWith = ClassFactory::InitListItem(itemIndex, this);
		}
	}
	SearchModeConstants searchMode = static_cast<SearchModeConstants>(0);
	SearchDirectionConstants searchDirection = static_cast<SearchDirectionConstants>(0);
	VARIANT searchFor;
	VariantInit(&searchFor);

	if(pDetails->lvfi.flags & LVFI_STRING) {
		searchMode = smText;
		searchFor.vt = VT_BSTR;
		if(pNotificationDetails->code == LVN_ODFINDITEMA) {
			searchFor.bstrVal = CComBSTR(reinterpret_cast<LPNMLVFINDITEMA>(pDetails)->lvfi.psz);
		} else {
			searchFor.bstrVal = CComBSTR(reinterpret_cast<LPNMLVFINDITEMW>(pDetails)->lvfi.psz);
		}
	} else if(pDetails->lvfi.flags & LVFI_PARTIAL) {
		searchMode = smPartialText;
		searchFor.vt = VT_BSTR;
		if(pNotificationDetails->code == LVN_ODFINDITEMA) {
			searchFor.bstrVal = CComBSTR(reinterpret_cast<LPNMLVFINDITEMA>(pDetails)->lvfi.psz);
		} else {
			searchFor.bstrVal = CComBSTR(reinterpret_cast<LPNMLVFINDITEMW>(pDetails)->lvfi.psz);
		}
	} else if(pDetails->lvfi.flags & LVFI_NEARESTXY) {
		searchMode = smNearestPosition;

		CComSafeArray<LONG> position;
		position.Create(2);
		position.SetAt(0, pDetails->lvfi.pt.x);
		position.SetAt(1, pDetails->lvfi.pt.y);
		searchFor.parray = SafeArrayCreateVectorEx(VT_I4, 0, 2, NULL);
		position.CopyTo(&searchFor.parray);
		searchFor.vt = VT_ARRAY | VT_I4;
		searchDirection = static_cast<SearchDirectionConstants>(pDetails->lvfi.vkDirection);
	} else if(pDetails->lvfi.flags & LVFI_PARAM) {
		searchMode = smItemData;
		searchFor.vt = VT_I4;
		searchFor.lVal = static_cast<LONG>(pDetails->lvfi.lParam);
	}

	IListViewItem* pFoundItem = NULL;
	Raise_FindVirtualItem(pItemToStartWith, searchMode, &searchFor, searchDirection, BOOL2VARIANTBOOL((pDetails->lvfi.flags & LVFI_WRAP) == LVFI_WRAP), &pFoundItem);

	VariantClear(&searchFor);
	if(pFoundItem) {
		LONG i = -1;
		pFoundItem->get_Index(&i);
		pFoundItem->Release();
		return i;
	}
	return -1;
}

LRESULT ExplorerListView::OnODStateChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVODSTATECHANGE pDetails = reinterpret_cast<LPNMLVODSTATECHANGE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLVODSTATECHANGE);
	if(!pDetails) {
		return 0;
	}

	LVITEMINDEX firstItemIndex = {pDetails->iFrom, 0};
	LVITEMINDEX lastItemIndex = {pDetails->iTo, 0};
	CComPtr<IListViewItem> pFirstLvwItem = ClassFactory::InitListItem(firstItemIndex, this);
	CComPtr<IListViewItem> pLastLvwItem = ClassFactory::InitListItem(lastItemIndex, this);
	if(!(pDetails->uOldState & LVIS_SELECTED) && (pDetails->uNewState & LVIS_SELECTED) == LVIS_SELECTED) {
		Raise_SelectedItemRange(pFirstLvwItem, pLastLvwItem);
	}

	return 0;
}

LRESULT ExplorerListView::OnSetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLVDISPINFO* pDetails = reinterpret_cast<NMLVDISPINFO*>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == LVN_SETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMLVDISPINFOW);
		}
	#endif

	// according to Windows SDK, we don't need to check pDetails->item.mask
	LVITEMINDEX itemIndex = {pDetails->item.iItem, pDetails->item.iGroup};
	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
	CComBSTR itemText;
	if(pNotificationDetails->code == LVN_SETDISPINFOA) {
		itemText = reinterpret_cast<LPNMLVDISPINFOA>(pDetails)->item.pszText;
	} else {
		itemText = reinterpret_cast<LPNMLVDISPINFOW>(pDetails)->item.pszText;
	}
	Raise_ItemSetText(pLvwItem, itemText);
	return 0;
}

LRESULT ExplorerListView::OnHeaderBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	if((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_HEADERDRAGDROP) == 0) {
		// the client doesn't want drag'n'drop
		return TRUE;
	}
	if(dragDropStatus.HeaderIsDragging()) {
		// we're already within a drag'n'drop operation
		return TRUE;
	}

	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMHEADER);
	if(!pDetails) {
		wasHandled = FALSE;
		return FALSE;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	pDetails->iButton = 1/*MouseButtonConstants.vbLeftButton*/;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	containedSysHeader32.ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this);

	if(pDetails->iButton == 1) {
		VARIANT_BOOL automaticDragDrop = VARIANT_TRUE;
		Raise_ColumnBeginDrag(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails), &automaticDragDrop);
		if(automaticDragDrop == VARIANT_FALSE) {
			return TRUE;
		}
		dragDropStatus.draggedColumn = pDetails->iItem;
	} else {
		ATLASSERT(pDetails->iButton == 1);
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnHeaderBeginTrackNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == HDN_BEGINTRACKA) {
			ATLASSERT_POINTER(pDetails, NMHEADERA);
		} else {
			ATLASSERT_POINTER(pDetails, NMHEADERW);
		}
	#endif

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_BeginColumnResizing(pLvwColumn, &cancel);
	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnHeaderDividerDblClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == HDN_DIVIDERDBLCLICKA) {
			ATLASSERT_POINTER(pDetails, NMHEADERA);
		} else {
			ATLASSERT_POINTER(pDetails, NMHEADERW);
		}
	#endif

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
	VARIANT_BOOL autoSize = VARIANT_TRUE;
	Raise_HeaderDividerDblClick(pLvwColumn, &autoSize);
	if(autoSize == VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnHeaderEndDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	if((SendMessage(LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0) & LVS_EX_HEADERDRAGDROP) == 0) {
		return TRUE;
	}

	// otherwise Raise_HeaderMouseUp would raise the HeaderClick event
	mouseStatus_Header.RemoveAllClickCandidates();

	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMHEADER);
	if(!pDetails) {
		wasHandled = FALSE;
		return FALSE;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	pDetails->iButton = 1/*MouseButtonConstants.vbLeftButton*/;
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	containedSysHeader32.ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this);

	KillTimer(timers.ID_DRAGSCROLL);
	if(pDetails->iButton == 1) {
		VARIANT_BOOL automaticDrop = VARIANT_TRUE;
		Raise_ColumnEndAutoDragDrop(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails), &automaticDrop);
		dragDropStatus.draggedColumn = -1;
		if(automaticDrop == VARIANT_FALSE) {
			return TRUE;
		}
	} else {
		ATLASSERT(pDetails->iButton == 1);
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnHeaderEndTrackNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == HDN_ENDTRACKA) {
			ATLASSERT_POINTER(pDetails, NMHEADERA);
		} else {
			ATLASSERT_POINTER(pDetails, NMHEADERW);
		}
	#endif

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
	Raise_EndColumnResizing(pLvwColumn);
	return 0;
}

LRESULT ExplorerListView::OnHeaderFilterBtnClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMHDFILTERBTNCLICK pDetails = reinterpret_cast<LPNMHDFILTERBTNCLICK>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMHDFILTERBTNCLICK);
	if(!pDetails) {
		wasHandled = FALSE;
		return FALSE;
	}

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	containedSysHeader32.ScreenToClient(&mousePosition);

	VARIANT_BOOL raiseFilterChanged = VARIANT_TRUE;
	Raise_FilterButtonClick(pLvwColumn, button, shift, mousePosition.x, mousePosition.y, reinterpret_cast<RECTANGLE*>(&pDetails->rc), &raiseFilterChanged);
	if(raiseFilterChanged != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnHeaderFilterChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	wasHandled = FALSE;

	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMHEADER);
	if(!pDetails) {
		return 0;
	}

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this);
	Raise_FilterChanged(pLvwColumn);
	return 0;
}

LRESULT ExplorerListView::OnHeaderGetDispInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMHDDISPINFO pDetails = reinterpret_cast<LPNMHDDISPINFO>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == HDN_GETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMHDDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMHDDISPINFOW);
		}
	#endif

	/* Very ugly HACK: pDetails->iItem often contains garbage. We try to detect such cases and try to
	   find the correct item by using the pDetails->lParam. Therefore the client should ensure it uses
	   unique lParam values, if it wants to make use of HDN_GETDISPINFO. More details:
	   https://groups.google.de/group/microsoft.public.win32.programmer.ui/browse_frm/thread/79dbd19dfa736b2e */
	BOOL brokenIndex = FALSE;
	HDITEM column = {0};
	column.mask = HDI_LPARAM;
	if(::SendMessage(pDetails->hdr.hwndFrom, HDM_GETITEM, pDetails->iItem, reinterpret_cast<LPARAM>(&column))) {
		brokenIndex = (pDetails->lParam != column.lParam);
	} else {
		brokenIndex = TRUE;
	}
	if(brokenIndex) {
		// iterate all columns and find the matching lParam
		int columns = static_cast<int>(::SendMessage(pDetails->hdr.hwndFrom, HDM_GETITEMCOUNT, 0, 0));
		int i = 0;
		for(i = 0; i < columns; ++i) {
			column.lParam = 0;
			::SendMessage(pDetails->hdr.hwndFrom, HDM_GETITEM, i, reinterpret_cast<LPARAM>(&column));
			if(column.lParam == pDetails->lParam) {
				pDetails->iItem = i;
				break;
			}
		}
		if(i == columns) {
			// no success - abort
			return 0;
		}
	}

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
	LONG iconIndex = 0;
	CComBSTR columnCaption = L"";
	LONG maxColumnCaptionLength = 0;

	RequestedInfoConstants requestedInfo = static_cast<RequestedInfoConstants>(0);
	if(pDetails->mask & HDI_IMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIconIndex);
	}
	if(pDetails->mask & HDI_TEXT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riItemText);
		// exclude the terminating 0
		maxColumnCaptionLength = pDetails->cchTextMax - 1;
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(pNotificationDetails->code == HDN_GETDISPINFOA) {
				columnCaption = reinterpret_cast<LPNMHDDISPINFOA>(pDetails)->pszText;
			} else {
				columnCaption = reinterpret_cast<LPNMHDDISPINFOW>(pDetails)->pszText;
			}
		#endif
	}

	VARIANT_BOOL dontAskAgain = VARIANT_FALSE;
	Raise_HeaderItemGetDisplayInfo(pLvwColumn, requestedInfo, &iconIndex, maxColumnCaptionLength, &columnCaption, &dontAskAgain);

	if(pDetails->mask & HDI_IMAGE) {
		pDetails->iImage = iconIndex;
	}
	if(pDetails->mask & HDI_TEXT) {
		// don't forget the terminating 0
		int bufferSize = columnCaption.Length() + 1;
		if(bufferSize > pDetails->cchTextMax) {
			bufferSize = pDetails->cchTextMax;
		}
		if(pNotificationDetails->code == HDN_GETDISPINFOA) {
			lstrcpynA(reinterpret_cast<LPNMHDDISPINFOA>(pDetails)->pszText, CW2A(columnCaption), bufferSize);
		} else {
			lstrcpynW(reinterpret_cast<LPNMHDDISPINFOW>(pDetails)->pszText, columnCaption, bufferSize);
		}
	}

	if(dontAskAgain == VARIANT_FALSE) {
		pDetails->mask &= ~HDI_DI_SETITEM;
	} else {
		pDetails->mask |= HDI_DI_SETITEM;
	}

	return 0;
}

LRESULT ExplorerListView::OnHeaderItemChangingNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == HDN_ITEMCHANGINGA) {
			ATLASSERT_POINTER(pDetails, NMHEADERA);
		} else {
			ATLASSERT_POINTER(pDetails, NMHEADERW);
		}
	#endif

	if(pDetails->pitem->mask & HDI_WIDTH) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this);
		VARIANT_BOOL abort = VARIANT_FALSE;
		Raise_ResizingColumn(pLvwColumn, reinterpret_cast<PLONG>(&pDetails->pitem->cxy), &abort);
		if(abort != VARIANT_FALSE) {
			return TRUE;
		}
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ExplorerListView::OnHeaderItemDblClkNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deHeaderClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		button |= 1/*MouseButtonConstants.vbLeftButton*/;
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		containedSysHeader32.ScreenToClient(&mousePosition);
		Raise_HeaderDblClick(button, shift, mousePosition.x, mousePosition.y);
	}
	return 0;
}

LRESULT ExplorerListView::OnHeaderStateIconClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMHEADERW pDetails = reinterpret_cast<LPNMHEADERW>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMHEADERW);
	if(!pDetails) {
		return 0;
	}

	HDITEM column = {0};
	column.mask = HDI_FORMAT;
	if(containedSysHeader32.SendMessage(HDM_GETITEM, pDetails->iItem, reinterpret_cast<LPARAM>(&column))) {
		if(column.fmt & HDF_CHECKBOX) {
			LONG previousStateImage = 0;
			previousStateImage = ((column.fmt & HDF_CHECKED) == HDF_CHECKED ? 2 : 1);
			
			int c = ImageList_GetImageCount(cachedSettings.hHeaderStateImageList);
			// calc the new state image
			LONG newStateImage = (previousStateImage + 1) % (c + 1);
			if(newStateImage == 0) {
				// state images are one-based
				newStateImage = 1;
			}

			CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_ColumnStateImageChanging(pLvwColumn, previousStateImage, &newStateImage, siccbMouse, &cancel);
			if(newStateImage == previousStateImage) {
				cancel = VARIANT_TRUE;
			}
			if(cancel == VARIANT_FALSE) {
				if(newStateImage == 0) {
					column.fmt &= ~(HDF_CHECKBOX | HDF_CHECKED);
				} else {
					column.fmt |= HDF_CHECKBOX;
					switch(newStateImage) {
						case 1:
							column.fmt &= ~HDF_CHECKED;
							break;
						case 2:
							column.fmt |= HDF_CHECKED;
							break;
					}
				}
				if(containedSysHeader32.SendMessage(HDM_SETITEM, pDetails->iItem, reinterpret_cast<LPARAM>(&column))) {
					Raise_ColumnStateImageChanged(pLvwColumn, previousStateImage, newStateImage, siccbMouse);
				}
			}
		}
	}
	return 0;
}

//LRESULT ExplorerListView::OnHeaderOverflowClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
//{
//	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
//	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
//	SHORT button = 0;
//	SHORT shift = 0;
//	WPARAM2BUTTONANDSHIFT(-1, button, shift);
//	button |= 1/*MouseButtonConstants.vbLeftButton*/;
//	DWORD position = GetMessagePos();
//	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
//	containedSysHeader32.ScreenToClient(&mousePosition);
//	Raise_HeaderChevronClick(pLvwColumn, button, shift, mousePosition.x, mousePosition.y);
//	return 0;
//}

LRESULT ExplorerListView::OnHeaderTrackNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMHEADER pDetails = reinterpret_cast<LPNMHEADER>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == HDN_TRACKA) {
			ATLASSERT_POINTER(pDetails, NMHEADERA);
		} else {
			ATLASSERT_POINTER(pDetails, NMHEADERW);
		}
	#endif

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(pDetails->iItem, this, FALSE);
	VARIANT_BOOL abort = VARIANT_FALSE;
	Raise_ResizingColumn(pLvwColumn, reinterpret_cast<PLONG>(&pDetails->pitem->cxy), &abort);
	if(abort != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}


LRESULT ExplorerListView::OnEditChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	Raise_EditChange();
	return 0;
}


LRESULT ExplorerListView::OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMLVCUSTOMDRAW pDetails = reinterpret_cast<LPNMLVCUSTOMDRAW>(lParam);
	CustomDrawReturnValuesConstants returnValue = static_cast<CustomDrawReturnValuesConstants>(this->DefWindowProc(message, wParam, lParam));

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			returnValue = static_cast<CustomDrawReturnValuesConstants>(shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_CUSTOMDRAW, wParam, lParam));
		}
	#endif

	if(!(properties.disabledEvents & deCustomDraw)) {
		DWORD itemType = LVCDI_ITEM;
		if(RunTimeHelper::IsCommCtrl6()) {
			// dwItemType sometimes contains rubbish, but maybe LVCDI_ITEM, LVCDI_GROUP and LVCDI_ITEMSLIST are not the only LVCDI_* values
			ATLASSERT(!(pDetails->dwItemType > LVCDI_ITEMSLIST && pDetails->dwItemType <= 10));
			if(pDetails->dwItemType == LVCDI_ITEM || pDetails->dwItemType == LVCDI_GROUP || pDetails->dwItemType == LVCDI_ITEMSLIST) {
				itemType = pDetails->dwItemType;
			}
			//if(itemType == LVCDI_ITEM) {
				// possible values for clrFace: 24
				// possible values for iIconEffect: 478
				// possible values for iIconPhase: 24, 60, 80
				// possible values for iPartId: 480
				// possible values for iStateId: 79, 99, 164, 195
				/*ATLASSERT(pDetails->iIconEffect == 0 || pDetails->iIconEffect == 478);
				ATLASSERT(pDetails->iIconPhase == 0 || pDetails->iIconPhase == 24 || pDetails->iIconPhase == 60 || pDetails->iIconPhase == 80);
				ATLASSERT(pDetails->iPartId == 0 || pDetails->iPartId == 480);
				ATLASSERT(pDetails->iStateId == 0 || pDetails->iStateId == 79 || pDetails->iStateId == 99 || pDetails->iStateId == 164 || pDetails->iStateId == 195);*/
			//}
		}

		if(itemType == LVCDI_ITEM || itemType == LVCDI_ITEMSLIST) {
			// we're drawing an item
			CComPtr<IListViewItem> pLvwItem = NULL;
			CComPtr<IListViewSubItem> pLvwSubItem = NULL;
			OLE_COLOR colorText = 0;
			OLE_COLOR colorTextBackground = 0;
			if(pDetails->nmcd.dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
				LVITEMINDEX itemIndex = {static_cast<int>(pDetails->nmcd.dwItemSpec), 0};
				if(pDetails->nmcd.dwDrawStage & CDDS_ITEM) {
					pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
				}
				if(pDetails->nmcd.dwDrawStage & CDDS_SUBITEM) {
					if(!pLvwItem) {
						pLvwItem = ClassFactory::InitListItem(itemIndex, this, FALSE);
					}
					pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, pDetails->iSubItem, this, FALSE, FALSE);
				}
				colorText = pDetails->clrText;
				colorTextBackground = pDetails->clrTextBk;
				if(colorTextBackground == -1) {
					colorTextBackground = 0x80000000 | COLOR_WINDOW;
				}
			}

			Raise_CustomDraw(pLvwItem, pLvwSubItem, BOOL2VARIANTBOOL(itemType == LVCDI_ITEMSLIST), &colorText, &colorTextBackground, static_cast<CustomDrawStageConstants>(pDetails->nmcd.dwDrawStage), static_cast<CustomDrawItemStateConstants>(pDetails->nmcd.uItemState), HandleToLong(pDetails->nmcd.hdc), reinterpret_cast<RECTANGLE*>(&pDetails->nmcd.rc), &returnValue);

			if(pDetails->nmcd.dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
				pDetails->clrText = colorText;
				pDetails->clrTextBk = colorTextBackground;
			}
		} else if(itemType == LVCDI_GROUP) {
			// we're drawing a group
			OLE_COLOR colorText = pDetails->clrText;
			AlignmentConstants headerAlignment = alLeft;
			switch(pDetails->uAlign & (LVGA_HEADER_LEFT | LVGA_HEADER_CENTER | LVGA_HEADER_RIGHT)) {
				case LVGA_HEADER_CENTER:
					headerAlignment = alCenter;
					break;
				case LVGA_HEADER_RIGHT:
					headerAlignment = alRight;
					break;
			}
			AlignmentConstants footerAlignment = alLeft;
			switch(pDetails->uAlign & (LVGA_FOOTER_LEFT | LVGA_FOOTER_CENTER | LVGA_FOOTER_RIGHT)) {
				case LVGA_FOOTER_CENTER:
					footerAlignment = alCenter;
					break;
				case LVGA_FOOTER_RIGHT:
					footerAlignment = alRight;
					break;
			}

			CComPtr<IListViewGroup> pLvwGroup = ClassFactory::InitGroup(static_cast<int>(pDetails->nmcd.dwItemSpec), this, FALSE);
			Raise_GroupCustomDraw(pLvwGroup, &colorText, &headerAlignment, &footerAlignment, static_cast<CustomDrawStageConstants>(pDetails->nmcd.dwDrawStage), static_cast<CustomDrawItemStateConstants>(pDetails->nmcd.uItemState), HandleToLong(pDetails->nmcd.hdc), reinterpret_cast<RECTANGLE*>(&pDetails->nmcd.rc), reinterpret_cast<RECTANGLE*>(&pDetails->rcText), &returnValue);

			pDetails->uAlign &= ~(LVGA_HEADER_LEFT | LVGA_HEADER_CENTER | LVGA_HEADER_RIGHT);
			switch(headerAlignment) {
				case alLeft:
					pDetails->uAlign |= LVGA_HEADER_LEFT;
					break;
				case alCenter:
					pDetails->uAlign |= LVGA_HEADER_CENTER;
					break;
				case alRight:
					pDetails->uAlign |= LVGA_HEADER_RIGHT;
					break;
			}
			pDetails->uAlign &= ~(LVGA_FOOTER_LEFT | LVGA_FOOTER_CENTER | LVGA_FOOTER_RIGHT);
			switch(footerAlignment) {
				case alLeft:
					pDetails->uAlign |= LVGA_FOOTER_LEFT;
					break;
				case alCenter:
					pDetails->uAlign |= LVGA_FOOTER_CENTER;
					break;
				case alRight:
					pDetails->uAlign |= LVGA_FOOTER_RIGHT;
					break;
			}
			pDetails->clrText = colorText;
		}
	}

	return static_cast<LRESULT>(returnValue);
}

LRESULT ExplorerListView::OnHeaderCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMCUSTOMDRAW pDetails = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
	// NOTE: This notification was forwarded to the control's parent.
	CustomDrawReturnValuesConstants returnValue = static_cast<CustomDrawReturnValuesConstants>(this->DefWindowProc(message, wParam, lParam));

	if(!(properties.disabledEvents & deHeaderCustomDraw)) {
		CComPtr<IListViewColumn> pLvwColumn = NULL;
		if((pDetails->dwDrawStage & CDDS_ITEM) || (pDetails->dwDrawStage & CDDS_SUBITEM)) {
			pLvwColumn = ClassFactory::InitColumn(static_cast<int>(pDetails->dwItemSpec), this, FALSE);
		}
		Raise_HeaderCustomDraw(pLvwColumn, static_cast<CustomDrawStageConstants>(pDetails->dwDrawStage), static_cast<CustomDrawItemStateConstants>(pDetails->uItemState), HandleToLong(pDetails->hdc), reinterpret_cast<RECTANGLE*>(&pDetails->rc), &returnValue);
	}

	return static_cast<LRESULT>(returnValue);
}

LRESULT ExplorerListView::OnToolTipGetDispInfoNotificationA(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	NMTTDISPINFOA backup = *reinterpret_cast<LPNMTTDISPINFOA>(lParam);
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(flags.cancelToolTip) {
		flags.cancelToolTip = FALSE;
		/* TODO: Find a better way than copying the old values back. Don't we produce a memleak here?
		         Who frees the string buffer (reinterpret_cast<LPNMTTDISPINFO>(lParam)->lpszText) that the
		         listview set up? */
		*reinterpret_cast<LPNMTTDISPINFOA>(lParam) = backup;
		lr = 0;     // actually the return value is ignored
	}

	return lr;
}

LRESULT ExplorerListView::OnToolTipGetDispInfoNotificationW(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	NMTTDISPINFOW backup = *reinterpret_cast<LPNMTTDISPINFOW>(lParam);
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(flags.cancelToolTip) {
		flags.cancelToolTip = FALSE;
		/* TODO: Find a better way than copying the old values back. Don't we produce a memleak here?
		         Who frees the string buffer (reinterpret_cast<LPNMTTDISPINFO>(lParam)->lpszText) that the
		         listview set up? */
		*reinterpret_cast<LPNMTTDISPINFOW>(lParam) = backup;
		lr = 0;     // actually the return value is ignored
	}

	return lr;
}


void ExplorerListView::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// retrieve the system font
			LOGFONT iconFont = {0};
			SystemParametersInfo(SPI_GETICONTITLELOGFONT, sizeof(LOGFONT), &iconFont, 0);
			properties.font.currentFont.CreateFontIndirect(&iconFont);
			properties.font.owningFontResource = TRUE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so just update the listview's font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT ExplorerListView::Raise_AbortedDrag(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AbortedDrag();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_AfterScroll(LONG dx, LONG dy)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AfterScroll(dx, dy);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_BeforeScroll(LONG dx, LONG dy)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeScroll(dx, dy);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_BeginColumnResizing(IListViewColumn* pColumn, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginColumnResizing(pColumn, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_BeginMarqueeSelection(VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginMarqueeSelection(pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CacheItemsHint(IListViewItem* pFirstItem, IListViewItem* pLastItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CacheItemsHint(pFirstItem, pLastItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CancelSubItemEdit(IListViewSubItem* pListSubItem, SubItemEditModeConstants editMode)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CancelSubItemEdit(pListSubItem, editMode);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CaretChanged(LVITEMINDEX previousCaretItem, LVITEMINDEX newCaretItem)
{
	// NOTE: If we ever pass this method's parameters by reference, the next 3 lines must be moved to the end of the method.
	caretChangedStatus.previousCaretItem.iItem = -1;
	caretChangedStatus.previousCaretItem.iGroup = 0;
	caretChangedStatus.newCaretItem = caretChangedStatus.previousCaretItem;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwPreviousCaretItem = ClassFactory::InitListItem(previousCaretItem, this);
		CComPtr<IListViewItem> pLvwNewCaretItem = ClassFactory::InitListItem(newCaretItem, this);
		return Fire_CaretChanged(pLvwPreviousCaretItem, pLvwNewCaretItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ChangedSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedSortOrder(previousSortOrder, newSortOrder);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ChangedWorkAreas(IListViewWorkAreas* pWorkAreas)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangedWorkAreas(pWorkAreas);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ChangingSortOrder(SortOrderConstants previousSortOrder, SortOrderConstants newSortOrder, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangingSortOrder(previousSortOrder, newSortOrder, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ChangingWorkAreas(IVirtualListViewWorkAreas* pWorkAreas, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChangingWorkAreas(pWorkAreas, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &mouseStatus.lastClickedItem, &subItemIndex);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(mouseStatus.lastClickedItem, subItemIndex, this, FALSE);
		}
		return Fire_Click(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, mouseStatus.lastClickedItem.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CollapsedGroup(IListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CollapsedGroup(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnBeginDrag(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails, VARIANT_BOOL* pDoAutomaticDragDrop)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnBeginDrag(pColumn, button, shift, x, y, hitTestDetails, pDoAutomaticDragDrop);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnClick(pColumn, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnDropDown(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnDropDown(pColumn, button, shift, x, y, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnEndAutoDragDrop(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails, VARIANT_BOOL* pDoAutomaticDrop)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnEndAutoDragDrop(pColumn, button, shift, x, y, hitTestDetails, pDoAutomaticDrop);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnMouseEnter(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus_Header.enteredControl) {
		return Fire_ColumnMouseEnter(pColumn, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnMouseLeave(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HeaderHitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus_Header.enteredControl) {
		return Fire_ColumnMouseLeave(pColumn, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnStateImageChanged(IListViewColumn* pColumn, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnStateImageChanged(pColumn, previousStateImageIndex, newStateImageIndex, causedBy);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ColumnStateImageChanging(IListViewColumn* pColumn, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ColumnStateImageChanging(pColumn, previousStateImageIndex, pNewStateImageIndex, causedBy, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CompareGroups(IListViewGroup* pFirstGroup, IListViewGroup* pSecondGroup, CompareResultConstants* pResult)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CompareGroups(pFirstGroup, pSecondGroup, pResult);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CompareItems(IListViewItem* pFirstItem, IListViewItem* pSecondItem, CompareResultConstants* pResult)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CompareItems(pFirstItem, pSecondItem, pResult);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ConfigureSubItemControl(IListViewSubItem* pListSubItem, SubItemControlKindConstants controlKind, SubItemEditModeConstants editMode, SubItemControlConstants subItemControl, BSTR* pThemeAppName, BSTR* pThemeIDList, HFONT* phFont, COLORREF* pTextColor, IUnknown** ppPropertyDescription, PROPVARIANT* pPropertyValue)
{
	//if(m_nFreezeEvents == 0) {
		OLE_COLOR oleTextColor = *reinterpret_cast<OLE_COLOR*>(pTextColor);
		HRESULT hr = Fire_ConfigureSubItemControl(pListSubItem, controlKind, editMode, subItemControl, pThemeAppName, pThemeIDList, reinterpret_cast<LONG*>(phFont), &oleTextColor, ppPropertyDescription, *reinterpret_cast<LONG*>(&pPropertyValue));
		*pTextColor = OLECOLOR2COLORREF(oleTextColor);
		return hr;
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		#ifdef INCLUDESHELLBROWSERINTERFACE
			POINT menuPosition = {x, y};
		#endif
		if(x == -1 && y == -1) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// retrieve the caret item and propose its rectangle's middle as the menu's position
				LVITEMINDEX caretItem = {-1, 0};
				if(IsComctl32Version610OrNewer()) {
					SendMessage(LVM_GETNEXTITEMINDEX, reinterpret_cast<WPARAM>(&caretItem), MAKELPARAM(LVNI_FOCUSED, 0));
				} else {
					caretItem.iItem = static_cast<int>(SendMessage(LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(LVNI_FOCUSED, 0)));
				}
				if(caretItem.iItem != -1) {
					WTL::CRect itemRectangle;
					itemRectangle.left = LVIR_LABEL;
					BOOL succeeded = FALSE;
					if(IsComctl32Version610OrNewer() && SendMessage(LVM_ISGROUPVIEWENABLED, 0, 0) && (GetStyle() & LVS_OWNERDATA) == LVS_OWNERDATA) {
						succeeded = static_cast<BOOL>(SendMessage(LVM_GETITEMINDEXRECT, reinterpret_cast<WPARAM>(&caretItem), reinterpret_cast<LPARAM>(&itemRectangle)));
					} else {
						succeeded = static_cast<BOOL>(SendMessage(LVM_GETITEMRECT, caretItem.iItem, reinterpret_cast<LPARAM>(&itemRectangle)));
					}
					if(succeeded) {
						WTL::CPoint centerPoint = itemRectangle.CenterPoint();
						x = centerPoint.x;
						y = centerPoint.y;
					}
				}
			} else {
				return S_OK;
			}
		}

		UINT hitTestDetails = 0;
		LVITEMINDEX itemIndex;
		int subItemIndex = -1;
		HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);

		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}
		VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
		HRESULT hr = Fire_ContextMenu(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1), &showDefaultMenu);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(SUCCEEDED(hr) && shellBrowserInterface.pInternalMessageListener) {
				if(showDefaultMenu != VARIANT_FALSE) {
					shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_CONTEXTMENU, FALSE, reinterpret_cast<LPARAM>(&menuPosition));
				}
			}
		#endif
		return hr;
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CreatedEditControlWindow(HWND hWndEdit)
{
	// initialize some values that are needed during label-editing
	labelEditStatus.Reset();
	hEditBackColorBrush = CreateSolidBrush(OLECOLOR2COLORREF(properties.editBackColor));

	// subclass the edit control
	containedEdit.SubclassWindow(hWndEdit);

	if(containedSysHeader32.IsWindow()) {
		if(containedSysHeader32.GetExStyle() & WS_EX_RTLREADING) {
			containedEdit.ModifyStyleEx(0, WS_EX_RTLREADING, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		} else {
			containedEdit.ModifyStyleEx(WS_EX_RTLREADING, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedEditControlWindow(HandleToLong(hWndEdit));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CreatedHeaderControlWindow(HWND hWndHeader)
{
	if(containedSysHeader32 != hWndHeader) {
		// subclass the header control
		containedSysHeader32.SubclassWindow(hWndHeader);
	}

	// configure the header control
	if(containedSysHeader32.IsWindow()) {
		if(properties.clickableColumnHeaders) {
			containedSysHeader32.ModifyStyle(0, HDS_BUTTONS, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		} else {
			containedSysHeader32.ModifyStyle(HDS_BUTTONS, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		}
		if(properties.headerFullDragging) {
			containedSysHeader32.ModifyStyle(0, HDS_FULLDRAG, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		} else {
			containedSysHeader32.ModifyStyle(HDS_FULLDRAG, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		}
		if(properties.headerHotTracking) {
			containedSysHeader32.ModifyStyle(0, HDS_HOTTRACK, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		} else {
			containedSysHeader32.ModifyStyle(HDS_HOTTRACK, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		}
		//if(IsComctl32Version580OrNewer()) {
			if(properties.showFilterBar) {
				containedSysHeader32.ModifyStyle(0, HDS_FILTERBAR, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			} else {
				containedSysHeader32.ModifyStyle(HDS_FILTERBAR, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			}
			if(properties.filterChangedTimeout != -1) {
				containedSysHeader32.SendMessage(HDM_SETFILTERCHANGETIMEOUT, 0, properties.filterChangedTimeout);
			}

			if(IsComctl32Version610OrNewer()) {
				if(properties.resizableColumns) {
					containedSysHeader32.ModifyStyle(HDS_NOSIZING, 0);
				} else {
					containedSysHeader32.ModifyStyle(0, HDS_NOSIZING);
				}
				if(properties.showHeaderChevron) {
					containedSysHeader32.ModifyStyle(0, HDS_OVERFLOW);
				} else {
					containedSysHeader32.ModifyStyle(HDS_OVERFLOW, 0);
				}
				if(properties.showHeaderStateImages) {
					containedSysHeader32.ModifyStyle(0, HDS_CHECKBOXES);
				} else {
					containedSysHeader32.ModifyStyle(HDS_CHECKBOXES, 0);
				}
			}
		//}

		/* HACK: SysListView32 seems to store the header's handle later, but we need it NOW (otherwise the
		         client won't be able to insert columns in the CreatedHeaderControlWindow event handler). So
		         tell the listview to store it now using a trick: You can set the listview's header to a custom
		         header control by passing its handle as lParam in LVM_GETHEADER. */
		SendMessage(LVM_GETHEADER, 0, reinterpret_cast<LPARAM>(containedSysHeader32.m_hWnd));

		if(GetExStyle() & WS_EX_RTLREADING) {
			containedSysHeader32.ModifyStyleEx(0, WS_EX_RTLREADING, SWP_DRAWFRAME | SWP_FRAMECHANGED);

			HDITEM column = {0};
			column.mask = HDI_FORMAT;
			int columns = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0));
			for(int columnIndex = 0; columnIndex < columns; ++columnIndex) {
				containedSysHeader32.SendMessage(HDM_GETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
				column.fmt |= HDF_RTLREADING;
				containedSysHeader32.SendMessage(HDM_SETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
			}
		} else {
			containedSysHeader32.ModifyStyleEx(WS_EX_RTLREADING, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);

			HDITEM column = {0};
			column.mask = HDI_FORMAT;
			int columns = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0));
			for(int columnIndex = 0; columnIndex < columns; ++columnIndex) {
				containedSysHeader32.SendMessage(HDM_GETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
				column.fmt &= ~HDF_RTLREADING;
				containedSysHeader32.SendMessage(HDM_SETITEM, columnIndex, reinterpret_cast<LPARAM>(&column));
			}
		}
		if(GetExStyle() & WS_EX_LAYOUTRTL) {
			containedSysHeader32.ModifyStyleEx(0, WS_EX_LAYOUTRTL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		} else {
			containedSysHeader32.ModifyStyleEx(WS_EX_LAYOUTRTL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		}

		// the header won't resize automatically, so force a resize by temporarily hiding it
		if(GetStyle() & LVS_NOCOLUMNHEADER) {
			ModifyStyle(LVS_NOCOLUMNHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			ModifyStyle(0, LVS_NOCOLUMNHEADER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		} else {
			ModifyStyle(0, LVS_NOCOLUMNHEADER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			ModifyStyle(LVS_NOCOLUMNHEADER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_CreatedHeaderControlWindow(HandleToLong(hWndHeader));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_CustomDraw(IListViewItem* pListItem, IListViewSubItem* pListSubItem, VARIANT_BOOL drawAllItems, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomDraw(pListItem, pListSubItem, drawAllItems, pTextColor, pTextBackColor, drawStage, itemState, hDC, pDrawingRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);
	if(itemIndex.iItem != mouseStatus.lastClickedItem.iItem || itemIndex.iGroup != mouseStatus.lastClickedItem.iGroup) {
		itemIndex.iItem = -1;
	}
	mouseStatus.lastClickedItem.iItem = -1;
	mouseStatus.lastClickedItem.iGroup = 0;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}
		return Fire_DblClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_DestroyedControlWindow(HWND hWnd)
{
	// clean up
	if(hBuiltInStateImageList) {
		/* We had a state imagelist. If it's still in use, we must destroy it.
		   NOTE: Never destroy a custom imagelist at this place. */
		if(hBuiltInStateImageList == reinterpret_cast<HIMAGELIST>(SendMessage(LVM_GETIMAGELIST, LVSIL_STATE, 0))) {
			ImageList_Destroy(hBuiltInStateImageList);
			if(cachedSettings.hStateImageList == hBuiltInStateImageList) {
				cachedSettings.hStateImageList = NULL;
			}
			hBuiltInStateImageList = NULL;
		}
	}
	// SysListView32 destroys the background bitmap, so keep us synchronized
	VARIANT v;
	VariantInit(&v);
	HBITMAP hPreviousBitmap = NULL;
	if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4))) {
		hPreviousBitmap = static_cast<HBITMAP>(LongToHandle(v.lVal));
	}
	if(properties.ownsBkImageBitmap && hPreviousBitmap && GetObjectType(hPreviousBitmap) == OBJ_BITMAP) {
		DeleteObject(hPreviousBitmap);
	}
	VariantClear(&properties.bkImage);

	KillTimer(timers.ID_CREATED);
	KillTimer(timers.ID_REDRAW);
	KillTimer(timers.ID_CARETCHANGED);
	RemoveItemFromItemContainers(-1);
	if(!RunTimeHelper::IsCommCtrl6()) {
		#ifdef USE_STL
			itemIDs.clear();
			itemParams.clear();
			groups.clear();
		#else
			itemIDs.RemoveAll();
			itemParams.RemoveAll();
			groups.RemoveAll();
		#endif
	}
	#ifdef USE_STL
		columnIndexes.clear();
		columnParams.clear();
	#else
		columnIndexes.RemoveAll();
		columnParams.RemoveAll();
	#endif

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_DestroyedEditControlWindow(HWND hWndEdit)
{
	BOOL acceptText = labelEditStatus.acceptText;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		LVITEMINDEX editedItem = labelEditStatus.editedItem;
		WCHAR pPreviousItemText[MAX_ITEMTEXTLENGTH + 1];
		if(labelEditStatus.previousText) {
			lstrcpynW(pPreviousItemText, OLE2W(labelEditStatus.previousText), MAX_ITEMTEXTLENGTH + 1);
		} else {
			pPreviousItemText[0] = L'\0';
		}
	#endif

	if(labelEditStatus.editedItem.iItem != -1) {
		if(acceptText) {
			CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(labelEditStatus.editedItem, this);
			BSTR itemText = NULL;
			pLvwItem->get_Text(&itemText);
			Raise_RenamedItem(pLvwItem, labelEditStatus.previousText, itemText);
			SysFreeString(itemText);
		}
		labelEditStatus.Reset();
	}

	if(!(properties.disabledEvents & deEditMouseEvents)) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedEdit;
		trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		if(mouseStatus_Edit.enteredControl) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			Raise_EditMouseLeave(button, shift, mouseStatus_Edit.lastPosition.x, mouseStatus_Edit.lastPosition.y);
		}
	}

	containedEdit.UnsubclassWindow();
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DestroyedEditControlWindow(HandleToLong(hWndEdit));
	//}

	// clean up
	DeleteObject(hEditBackColorBrush);
	hEditBackColorBrush = NULL;

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(!acceptText) {
			// the shell item could not be renamed, so go back into edit mode
			CComPtr<IListView_WIN7> pListView7 = NULL;
			CComPtr<IListView_WINVISTA> pListViewVista = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7)) && pListView7) {
				ATLASSUME(pListView7);
				ATLVERIFY(SUCCEEDED(pListView7->SetItemText(editedItem.iItem, 0, pPreviousItemText)));
			} else if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista)) && pListViewVista) {
				ATLASSUME(pListViewVista);
				ATLVERIFY(SUCCEEDED(pListViewVista->SetItemText(editedItem.iItem, 0, pPreviousItemText)));
			} else {
				TCHAR pBuffer[MAX_ITEMTEXTLENGTH + 1];
				CW2T converter(pPreviousItemText);
				lstrcpyn(pBuffer, converter, MAX_ITEMTEXTLENGTH + 1);
				LVITEM item = {0};
				item.pszText = pBuffer;
				item.cchTextMax = lstrlen(item.pszText);
				SendMessage(LVM_SETITEMTEXT, editedItem.iItem, reinterpret_cast<LPARAM>(&item));
			}
			PostMessage(LVM_EDITLABEL, editedItem.iItem, 0);
		}
	#endif

	return hr;
}

inline HRESULT ExplorerListView::Raise_DestroyedHeaderControlWindow(HWND hWndHeader)
{
	// clean up
	if(hBuiltInHeaderStateImageList) {
		/* We had a state imagelist. If it's still in use, we must destroy it.
		   NOTE: Never destroy a custom imagelist at this place. */
		if(hBuiltInHeaderStateImageList == reinterpret_cast<HIMAGELIST>(containedSysHeader32.SendMessage(HDM_GETIMAGELIST, HDSIL_STATE, 0))) {
			ImageList_Destroy(hBuiltInHeaderStateImageList);
			hBuiltInHeaderStateImageList = NULL;
		}
	}

	containedSysHeader32.UnsubclassWindow();
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedHeaderControlWindow(HandleToLong(hWndHeader));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_DragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(dragDropStatus.hDragImageList) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ImageList_DragMove(mousePosition.x, mousePosition.y);
	}

	UINT hitTestDetails = 0;
	LVITEMINDEX dropTarget;
	HitTest(x, y, &hitTestDetails, &dropTarget, NULL, TRUE);
	IListViewItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CPoint mousePosition(x, y);
		WTL::CRect noScrollZone(0, 0, 0, 0);
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePosition) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			if(containedSysHeader32.IsWindow() && containedSysHeader32.IsWindowVisible()) {
				WTL::CRect headerRectangle;
				containedSysHeader32.GetWindowRect(&headerRectangle);
				noScrollZone.top += headerRectangle.Height();
			}
			isInScrollZone = !noScrollZone.PtInRect(mousePosition);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePosition.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePosition.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePosition.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePosition.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_DragMouseMove(&pDropTarget, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1), &autoHScrollVelocity, &autoVScrollVelocity);
	//}

	if(pDropTarget) {
		// TODO: Retrieving just the index isn't enough!
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget.iItem = l;
		// we don't want to produce a mem-leak...
		pDropTarget->Release();
	} else {
		dropTarget.iItem = -1;
		dropTarget.iGroup = 0;
	}

	dragDropStatus.lastDropTarget = dropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() >> 2 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_Drop(void)
{
	//if(m_nFreezeEvents == 0) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);

		UINT hitTestDetails = 0;
		LVITEMINDEX dropTarget;
		HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &dropTarget, NULL, TRUE);
		CComPtr<IListViewItem> pDropTarget = ClassFactory::InitListItem(dropTarget, this);

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		return Fire_Drop(pDropTarget, button, shift, mousePosition.x, mousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditChange(void)
{
	if(!containedEdit.IsWindow()) {
		/* We ignore the event because the user would probably get/set the EditText property which
		   fails as long as the window isn't created. */
		return S_OK;
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_EditChange();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	if((x == -1) && (y == -1)) {
		// the event was caused by the keyboard
		if(properties.processContextMenuKeys) {
			// propose the middle of the edit control's client rectangle as the menu's position
			WTL::CRect clientArea;
			containedEdit.GetClientRect(&clientArea);
			WTL::CPoint centerPoint = clientArea.CenterPoint();
			x = centerPoint.x;
			y = centerPoint.y;
		} else {
			return S_OK;
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_EditContextMenu(button, shift, x, y, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditGotFocus(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditGotFocus();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditKeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditKeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditKeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditKeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditKeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditKeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditLostFocus(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditLostFocus();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Edit.enteredControl) {
		Raise_EditMouseEnter(button, shift, x, y);
	}
	if(!mouseStatus_Edit.hoveredControl) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = containedEdit.m_hWnd;
		trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		Raise_EditMouseHover(button, shift, x, y);
	}
	mouseStatus_Edit.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseDown(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = containedEdit.m_hWnd;
	trackingOptions.dwHoverTime = (properties.editHoverTime == -1 ? HOVER_DEFAULT : properties.editHoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus_Edit.EnterControl();
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseEnter(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_Edit.HoverControl();
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseHover(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus_Edit.LeaveControl();
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseLeave(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Edit.enteredControl) {
		Raise_EditMouseEnter(button, shift, x, y);
	}
	mouseStatus_Edit.lastPosition.x = x;
	mouseStatus_Edit.lastPosition.y = y;
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus_Edit.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		containedEdit.GetClientRect(&clientArea);
		containedEdit.ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != containedEdit.m_hWnd) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					Raise_EditClick(button, shift, x, y);
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					Raise_EditRClick(button, shift, x, y);
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					Raise_EditMClick(button, shift, x, y);
					break;
				case embXButton1:
				case embXButton2:
					Raise_EditXClick(button, shift, x, y);
					break;
			}
		}
		mouseStatus_Edit.RemoveClickCandidate(button);
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseUp(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus_Edit.enteredControl) {
		Raise_EditMouseEnter(button, shift, x, y);
	}
	mouseStatus_Edit.lastPosition.x = x;
	mouseStatus_Edit.lastPosition.y = y;
	//if(m_nFreezeEvents == 0) {
		return Fire_EditMouseWheel(button, shift, x, y, scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditRClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditRDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditRDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditXClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EditXDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EditXDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EmptyMarkupTextLinkClick(LONG linkIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EmptyMarkupTextLinkClick(linkIndex, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EndColumnResizing(IListViewColumn* pColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EndColumnResizing(pColumn);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_EndSubItemEdit(IListViewSubItem* pListSubItem, SubItemEditModeConstants editMode, PROPERTYKEY* pPropertyKey, PROPVARIANT* pPropertyValue, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_EndSubItemEdit(pListSubItem, editMode, *reinterpret_cast<LONG*>(&pPropertyKey), *reinterpret_cast<LONG*>(&pPropertyValue), pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ExpandedGroup(IListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ExpandedGroup(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FilterButtonClick(IListViewColumn* pColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, RECTANGLE* pFilterButtonRectangle, VARIANT_BOOL* pRaiseFilterChanged)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FilterButtonClick(pColumn, button, shift, x, y, pFilterButtonRectangle, pRaiseFilterChanged);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FilterChanged(IListViewColumn* pColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FilterChanged(pColumn);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FindVirtualItem(IListViewItem* pItemToStartWith, SearchModeConstants searchMode, VARIANT* pSearchFor, SearchDirectionConstants searchDirection, VARIANT_BOOL wrapAtLastItem, IListViewItem** ppFoundItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FindVirtualItem(pItemToStartWith, searchMode, pSearchFor, searchDirection, wrapAtLastItem, ppFoundItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FooterItemClick(IListViewFooterItem* pFooterItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pRemoveFooterArea)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FooterItemClick(pFooterItem, button, shift, x, y, hitTestDetails, pRemoveFooterArea);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FreeColumnData(IListViewColumn* pColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeColumnData(pColumn);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FreeFooterItemData(IListViewFooterItem* pFooterItem, LONG itemData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeFooterItemData(pFooterItem, itemData);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_FreeItemData(IListViewItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeItemData(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GetSubItemControl(IListViewSubItem* pListSubItem, SubItemControlKindConstants controlKind, SubItemEditModeConstants editMode, SubItemControlConstants* pSubItemControl)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetSubItemControl(pListSubItem, controlKind, editMode, pSubItemControl);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GroupAsynchronousDrawFailed(IListViewGroup* pGroup, NMLVASYNCDRAWN* pNotificationDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GroupAsynchronousDrawFailed(pGroup, pNotificationDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GroupCustomDraw(IListViewGroup* pGroup, OLE_COLOR* pTextColor, AlignmentConstants* pHeaderAlignment, AlignmentConstants* pFooterAlignment, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants groupState, LONG hDC, RECTANGLE* pDrawingRectangle, RECTANGLE* pTextRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GroupCustomDraw(pGroup, pTextColor, pHeaderAlignment, pFooterAlignment, drawStage, groupState, hDC, pDrawingRectangle, pTextRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GroupGotFocus(IListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GroupGotFocus(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GroupLostFocus(IListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GroupLostFocus(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GroupSelectionChanged(IListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GroupSelectionChanged(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_GroupTaskLinkClick(IListViewGroup* pGroup, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GroupTaskLinkClick(pGroup, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderAbortedDrag(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderAbortedDrag();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderChevronClick(IListViewColumn* pFirstOverflownColumn, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pShowDefaultMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderChevronClick(pFirstOverflownColumn, button, shift, x, y, pShowDefaultMenu);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus_Header.lastClickedItem = HeaderHitTest(x, y, &hitTestDetails);

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			if(hitTestDetails == HHT_ONHEADER) {
				LONG columnID = ColumnIndexToID(mouseStatus_Header.lastClickedItem);
				if(columnID != -1) {
					shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HEADERCLICK, columnID, 0);
				}
			}
		}
		if(properties.disabledEvents & deHeaderClickEvents) {
			return S_OK;
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(mouseStatus_Header.lastClickedItem, this);
		return Fire_HeaderClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		#ifdef INCLUDESHELLBROWSERINTERFACE
			POINT menuPosition = {x, y};
		#endif
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				// propose the middle of the header control's client rectangle as the menu's position
				WTL::CRect clientArea;
				containedSysHeader32.GetClientRect(&clientArea);
				WTL::CPoint centerPoint = clientArea.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
			} else {
				return S_OK;
			}
		}

		UINT hitTestDetails = 0;
		int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		VARIANT_BOOL showDefaultMenu = VARIANT_TRUE;
		HRESULT hr = Fire_HeaderContextMenu(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails), &showDefaultMenu);
		#ifdef INCLUDESHELLBROWSERINTERFACE
			if(SUCCEEDED(hr) && shellBrowserInterface.pInternalMessageListener) {
				if(showDefaultMenu != VARIANT_FALSE) {
					shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_CONTEXTMENU, TRUE, reinterpret_cast<LPARAM>(&menuPosition));
				}
			}
		#endif
		return hr;
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderCustomDraw(IListViewColumn* pColumn, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants columnState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderCustomDraw(pColumn, drawStage, columnState, hDC, pDrawingRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	flags.raisedHeaderDblClick = TRUE;

	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);
	if(columnIndex != mouseStatus_Header.lastClickedItem) {
		columnIndex = -1;
	}
	mouseStatus_Header.lastClickedItem = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		return Fire_HeaderDblClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderDividerDblClick(IListViewColumn* pColumn, VARIANT_BOOL* pAutoSizeColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderDividerDblClick(pColumn, pAutoSizeColumn);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderDragMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, OLE_XPOS_PIXELS xListView, OLE_YPOS_PIXELS yListView)
{
	if(dragDropStatus.hDragImageList) {
		if(dragDropStatus.restrictedDragImage) {
			RECT windowRect = {0};
			containedSysHeader32.GetWindowRect(&windowRect);
			ImageList_DragMove(GET_X_LPARAM(GetMessagePos()), windowRect.top);
		} else {
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ImageList_DragMove(mousePosition.x, mousePosition.y);
		}
	}

	UINT hitTestDetails = 0;
	int dropTarget = HeaderHitTest(x, 1, &hitTestDetails);
	if(dropTarget == -1) {
		if(hitTestDetails & hhtNoItem) {
			dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
		}
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CPoint mousePosition(xListView, yListView);
		WTL::CRect noScrollZone;
		// this will give us the complete client area (ignoring scroll position)
		containedSysHeader32.GetClientRect(&noScrollZone);
		// correct the client area
		WTL::CRect rc;
		GetClientRect(&rc);
		noScrollZone.left = rc.left;
		noScrollZone.right = rc.right;

		if(dragDropStatus.restrictedDragImage) {
			/* The header control's height is small and many users would have problems keeping the mouse cursor
			   within the header's client area, so Windows enlarges the height by a horizontal scrollbar's
			   height in each dimension and so do we. */
			NONCLIENTMETRICS nonClientMetrics = {0};
			nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, RunTimeHelper::SizeOf_NONCLIENTMETRICS(), &nonClientMetrics, 0);
			noScrollZone.InflateRect(0, nonClientMetrics.iScrollHeight * 2);
		}

		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePosition) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
			isInScrollZone = !noScrollZone.PtInRect(mousePosition);
		} else {
			// we're outside the (extended) client area
			dropTarget = -1;
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePosition.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePosition.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
		}
	}

	// we accept mouse cursor positions outside the client area, so calling HeaderHitTest once isn't enough
	hitTestDetails = 0;
	HeaderHitTest(x, y, &hitTestDetails);
	IListViewColumn* pDropTarget = NULL;
	ClassFactory::InitColumn(dropTarget, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_HeaderDragMouseMove(&pDropTarget, button, shift, x, y, xListView, yListView, static_cast<HeaderHitTestConstants>(hitTestDetails), &autoHScrollVelocity, &autoVScrollVelocity);
	//}

	if(pDropTarget) {
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget = l;
		// we don't want to produce a mem-leak...
		pDropTarget->Release();
	} else {
		dropTarget = -1;
	}

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() >> 2 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_HeaderDrop(void)
{
	//if(m_nFreezeEvents == 0) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		containedSysHeader32.ScreenToClient(&mousePosition);

		UINT hitTestDetails = 0;
		int dropTarget = HeaderHitTest(mousePosition.x, 1, &hitTestDetails);
		if(dropTarget == -1) {
			if(hitTestDetails & hhtNoItem) {
				dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
			}
		}
		// we accept mouse cursor positions outside the client area, so calling HeaderHitTest once isn't enough
		hitTestDetails = 0;
		HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

		CComPtr<IListViewColumn> pDropTarget = ClassFactory::InitColumn(dropTarget, this);

		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		return Fire_HeaderDrop(pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderGotFocus(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderGotFocus();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderItemGetDisplayInfo(IListViewColumn* pColumn, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG maxColumnCaptionLength, BSTR* pColumnCaption, VARIANT_BOOL* pDontAskAgain)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderItemGetDisplayInfo(pColumn, requestedInfo, pIconIndex, maxColumnCaptionLength, pColumnCaption, pDontAskAgain);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderKeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderKeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderKeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderKeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderKeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderKeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderLostFocus(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderLostFocus();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus_Header.lastClickedItem = HeaderHitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(mouseStatus_Header.lastClickedItem, this);
		return Fire_HeaderMClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);
	if(columnIndex != mouseStatus_Header.lastClickedItem) {
		columnIndex = -1;
	}
	mouseStatus_Header.lastClickedItem = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		return Fire_HeaderMDblClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

	if(!(properties.disabledEvents & deHeaderMouseEvents)) {
		if(!mouseStatus_Header.enteredControl) {
			Raise_HeaderMouseEnter(button, shift, x, y);
		}
		if(!mouseStatus_Header.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = containedSysHeader32.m_hWnd;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_HeaderMouseHover(button, shift, x, y);
		}
		mouseStatus_Header.StoreClickCandidate(button);
		mouseStatus_Header.mouseDownItem = columnIndex;

		//if(m_nFreezeEvents == 0) {
			CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
			return Fire_HeaderMouseDown(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
		//}
		//return S_OK;
	} else {
		if(!mouseStatus_Header.enteredControl) {
			mouseStatus_Header.EnterControl();
		}
		mouseStatus_Header.StoreClickCandidate(button);
		mouseStatus_Header.mouseDownItem = columnIndex;
		return S_OK;
	}
}

inline HRESULT ExplorerListView::Raise_HeaderMouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = containedSysHeader32.m_hWnd;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.headerHoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus_Header.EnterControl();

	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);
	columnUnderMouse = columnIndex;

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_HeaderMouseEnter(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	if(pLvwColumn) {
		Raise_ColumnMouseEnter(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	}
	return hr;
}

inline HRESULT ExplorerListView::Raise_HeaderMouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Header.hoveredControl) {
		mouseStatus_Header.HoverControl();

		UINT hitTestDetails = 0;
		int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

		//if(m_nFreezeEvents == 0) {
			CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
			return Fire_HeaderMouseHover(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
		//}
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnUnderMouse, this);
	if(pLvwColumn) {
		Raise_ColumnMouseLeave(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	}
	columnUnderMouse = -1;

	mouseStatus_Header.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		return Fire_HeaderMouseLeave(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus_Header.enteredControl) {
		Raise_HeaderMouseEnter(button, shift, x, y);
	}
	mouseStatus_Header.lastPosition.x = x;
	mouseStatus_Header.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);

	// Do we move over another column than before?
	if(columnIndex != columnUnderMouse) {
		CComPtr<IListViewColumn> pPrevLvwColumn = ClassFactory::InitColumn(columnUnderMouse, this);
		if(pPrevLvwColumn) {
			Raise_ColumnMouseLeave(pPrevLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
		}
		columnUnderMouse = columnIndex;
		if(pLvwColumn) {
			Raise_ColumnMouseEnter(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderMouseMove(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

	if(mouseStatus_Header.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		containedSysHeader32.GetClientRect(&clientArea);
		containedSysHeader32.ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != containedSysHeader32.m_hWnd) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					#ifdef INCLUDESHELLBROWSERINTERFACE
						//if(!(properties.disabledEvents & deHeaderClickEvents)) {
							if(columnIndex == mouseStatus_Header.mouseDownItem) {
								Raise_HeaderClick(button, shift, x, y);
							}
						//}
					#else
						if(!(properties.disabledEvents & deHeaderClickEvents)) {
							if(columnIndex == mouseStatus_Header.mouseDownItem) {
								Raise_HeaderClick(button, shift, x, y);
							}
						}
					#endif
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deHeaderClickEvents)) {
						if(columnIndex == mouseStatus_Header.mouseDownItem) {
							Raise_HeaderRClick(button, shift, x, y);
						}
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deHeaderClickEvents)) {
						if(columnIndex == mouseStatus_Header.mouseDownItem) {
							Raise_HeaderMClick(button, shift, x, y);
						}
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deHeaderClickEvents)) {
						if(columnIndex == mouseStatus_Header.mouseDownItem) {
							Raise_HeaderXClick(button, shift, x, y);
						}
					}
					break;
			}
		}

		mouseStatus_Header.mouseDownItem = -1;
		mouseStatus_Header.RemoveClickCandidate(button);
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deHeaderMouseEvents)) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		return Fire_HeaderMouseUp(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderMouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus_Header.enteredControl) {
		Raise_HeaderMouseEnter(button, shift, x, y);
	}
	mouseStatus_Header.lastPosition.x = x;
	mouseStatus_Header.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);

	CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderMouseWheel(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails), scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	POINT mousePositionListView = {mousePosition.x, mousePosition.y};
	containedSysHeader32.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	ScreenToClient(&mousePositionListView);
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	//UINT hitTestDetails = 0;
	//int dropTarget = HeaderHitTest(mousePosition.x, 1, &hitTestDetails);
	//if(dropTarget == -1) {
	//	if(hitTestDetails & hhtNoItem) {
	//		dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
	//	}
	//}
	//// we accept mouse cursor positions outside the client area, so calling HeaderHitTest once isn't enough
	//hitTestDetails = 0;
	//HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	UINT hitTestDetails = 0;
	int dropTarget = HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	if(dropTarget == -1) {
		if(hitTestDetails & hhtNoItem) {
			dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
		}
	}
	IListViewColumn* pDropTarget = NULL;
	ClassFactory::InitColumn(dropTarget, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_HeaderOLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = TRUE;
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.dropTarget = ColumnIndexToID(dropTarget);
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			*reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition) = mousePositionListView;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = static_cast<HeaderHitTestConstants>(hitTestDetails);
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DROP, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
					dragDropStatus.drop_pDataObject = NULL;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<IListViewColumn*>(eventDetails.pDropTarget);
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_HeaderOLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, mousePositionListView.x, mousePositionListView.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.HeaderOLEDragLeaveOrDrop();
	containedSysHeader32.Invalidate();

	return hr;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEDragEnter(BOOL fakedEnter, IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	POINT mousePositionListView = {mousePosition.x, mousePosition.y};
	containedSysHeader32.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	ScreenToClient(&mousePositionListView);
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.HeaderOLEDragEnter();

	//UINT hitTestDetails = 0;
	//int dropTarget = HeaderHitTest(mousePosition.x, 1, &hitTestDetails);
	//if(dropTarget == -1) {
	//	if(hitTestDetails & hhtNoItem) {
	//		dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
	//	}
	//}
	UINT hitTestDetails = 0;
	int dropTarget = HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	if(dropTarget == -1) {
		if(hitTestDetails & hhtNoItem) {
			dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
		}
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CRect noScrollZone;
		// this will give us the complete client area (ignoring scroll position)
		containedSysHeader32.GetClientRect(&noScrollZone);
		// correct the client area
		WTL::CRect rc;
		GetClientRect(&rc);
		noScrollZone.left = rc.left;
		noScrollZone.right = rc.right;

		///* The header control's height is small and many users would have problems keeping the mouse cursor
		//   within the header's client area, so Windows enlarges the height by a horizontal scrollbar's
		//   height in each dimension and so do we. */
		//NONCLIENTMETRICS nonClientMetrics = {0};
		//nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
		//SystemParametersInfo(SPI_GETNONCLIENTMETRICS, RunTimeHelper::SizeOf_NONCLIENTMETRICS(), &nonClientMetrics, 0);
		//noScrollZone.InflateRect(0, nonClientMetrics.iScrollHeight * 2);

		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePositionListView) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
			isInScrollZone = !noScrollZone.PtInRect(mousePositionListView);
		} else {
			// we're outside the (extended) client area
			dropTarget = -1;
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePositionListView.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePositionListView.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
		}
	}

	//// we accept mouse cursor positions outside the client area, so calling HeaderHitTest once isn't enough
	//hitTestDetails = 0;
	//HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	IListViewColumn* pDropTarget = NULL;
	ClassFactory::InitColumn(dropTarget, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(!fakedEnter) {
		if(pData) {
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
	}
	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = TRUE;
			eventDetails.pDataObject = (fakedEnter ? NULL : reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p));
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.dropTarget = ColumnIndexToID(dropTarget);
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			*reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition) = mousePositionListView;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = static_cast<HeaderHitTestConstants>(hitTestDetails);
			eventDetails.autoHScrollVelocity = autoHScrollVelocity;
			eventDetails.autoVScrollVelocity = autoVScrollVelocity;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGENTER, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<IListViewColumn*>(eventDetails.pDropTarget);
				autoHScrollVelocity = eventDetails.autoHScrollVelocity;
				autoVScrollVelocity = eventDetails.autoVScrollVelocity;
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject && !fakedEnter) {
			hr = Fire_HeaderOLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, mousePositionListView.x, mousePositionListView.y, static_cast<HeaderHitTestConstants>(hitTestDetails), &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(pDropTarget) {
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget = l;
		// we're using a raw pointer
		pDropTarget->Release();
	} else {
		dropTarget = -1;
	}
	dragDropStatus.lastDropTarget.iItem = dropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() >> 2 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEDragLeave(BOOL fakedLeave, BOOL* pCallDropTargetHelper)
{
	POINT mousePosition = {dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y};
	POINT mousePositionListView = mousePosition;
	containedSysHeader32.ScreenToClient(&mousePosition);
	ScreenToClient(&mousePositionListView);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	//UINT hitTestDetails = 0;
	//int dropTarget = HeaderHitTest(mousePosition.x, 1, &hitTestDetails);
	//if(dropTarget == -1) {
	//	if(hitTestDetails & hhtNoItem) {
	//		dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
	//	}
	//}
	//// we accept mouse cursor positions outside the client area, so calling HeaderHitTest once isn't enough
	//hitTestDetails = 0;
	//HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	UINT hitTestDetails = 0;
	int dropTarget = HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	if(dropTarget == -1) {
		if(hitTestDetails & hhtNoItem) {
			dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
		}
	}
	CComPtr<IListViewColumn> pDropTarget = ClassFactory::InitColumn(dropTarget, this);

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = TRUE;
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget.p);
			eventDetails.dropTarget = ColumnIndexToID(dropTarget);
			BUTTONANDSHIFT2OLEKEYSTATE(button, shift, eventDetails.keyState);
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			*reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition) = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = static_cast<HeaderHitTestConstants>(hitTestDetails);
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGLEAVE, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_HeaderOLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, mousePosition.x, mousePosition.y, mousePositionListView.x, mousePositionListView.y, static_cast<HeaderHitTestConstants>(hitTestDetails));
		}
	//}

	if(fakedLeave) {
		dragDropStatus.isOverHeader = FALSE;
	} else {
		dragDropStatus.pActiveDataObject = NULL;
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.HeaderOLEDragLeaveOrDrop();
		Invalidate();
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	dragDropStatus.lastMousePosition = mousePosition;
	POINT mousePositionListView = {mousePosition.x, mousePosition.y};
	containedSysHeader32.ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	ScreenToClient(&mousePositionListView);
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	//UINT hitTestDetails = 0;
	//int dropTarget = HeaderHitTest(mousePosition.x, 1, &hitTestDetails);
	//if(dropTarget == -1) {
	//	if(hitTestDetails & hhtNoItem) {
	//		dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
	//	}
	//}
	UINT hitTestDetails = 0;
	int dropTarget = HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	if(dropTarget == -1) {
		if(hitTestDetails & hhtNoItem) {
			dropTarget = static_cast<int>(containedSysHeader32.SendMessage(HDM_GETITEMCOUNT, 0, 0)) - 1;
		}
	}

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CRect noScrollZone;
		// this will give us the complete client area (ignoring scroll position)
		containedSysHeader32.GetClientRect(&noScrollZone);
		// correct the client area
		WTL::CRect rc;
		GetClientRect(&rc);
		noScrollZone.left = rc.left;
		noScrollZone.right = rc.right;

		///* The header control's height is small and many users would have problems keeping the mouse cursor
		//   within the header's client area, so Windows enlarges the height by a horizontal scrollbar's
		//   height in each dimension and so do we. */
		//NONCLIENTMETRICS nonClientMetrics = {0};
		//nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
		//SystemParametersInfo(SPI_GETNONCLIENTMETRICS, RunTimeHelper::SizeOf_NONCLIENTMETRICS(), &nonClientMetrics, 0);
		//noScrollZone.InflateRect(0, nonClientMetrics.iScrollHeight * 2);

		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePositionListView) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, 0, properties.DRAGSCROLLZONEWIDTH, 0);
			isInScrollZone = !noScrollZone.PtInRect(mousePositionListView);
		} else {
			// we're outside the (extended) client area
			dropTarget = -1;
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePositionListView.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePositionListView.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
		}
	}

	//// we accept mouse cursor positions outside the client area, so calling HeaderHitTest once isn't enough
	//hitTestDetails = 0;
	//HeaderHitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	IListViewColumn* pDropTarget = NULL;
	ClassFactory::InitColumn(dropTarget, this, IID_IListViewColumn, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = TRUE;
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.dropTarget = ColumnIndexToID(dropTarget);
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			*reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition) = mousePositionListView;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = static_cast<HeaderHitTestConstants>(hitTestDetails);
			eventDetails.autoHScrollVelocity = autoHScrollVelocity;
			eventDetails.autoVScrollVelocity = autoVScrollVelocity;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGOVER, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<IListViewColumn*>(eventDetails.pDropTarget);
				autoHScrollVelocity = eventDetails.autoHScrollVelocity;
				autoVScrollVelocity = eventDetails.autoVScrollVelocity;
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_HeaderOLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, mousePositionListView.x, mousePositionListView.y, static_cast<HeaderHitTestConstants>(hitTestDetails), &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(pDropTarget) {
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget = l;
		// we're using a raw pointer
		pDropTarget->Release();
	} else {
		dropTarget = -1;
	}
	dragDropStatus.lastDropTarget.iItem = dropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() >> 2 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_HeaderOLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ExplorerListView::Raise_HeaderOLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderOwnerDrawItem(IListViewColumn* pColumn, OwnerDrawItemStateConstants columnState, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeaderOwnerDrawItem(pColumn, columnState, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderRClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus_Header.lastClickedItem = HeaderHitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(mouseStatus_Header.lastClickedItem, this);
		return Fire_HeaderRClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderRDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);
	if(columnIndex != mouseStatus_Header.lastClickedItem) {
		columnIndex = -1;
	}
	mouseStatus_Header.lastClickedItem = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		return Fire_HeaderRDblClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderXClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus_Header.lastClickedItem = HeaderHitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(mouseStatus_Header.lastClickedItem, this);
		return Fire_HeaderXClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HeaderXDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int columnIndex = HeaderHitTest(x, y, &hitTestDetails);
	if(columnIndex != mouseStatus_Header.lastClickedItem) {
		columnIndex = -1;
	}
	mouseStatus_Header.lastClickedItem = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewColumn> pLvwColumn = ClassFactory::InitColumn(columnIndex, this);
		return Fire_HeaderXDblClick(pLvwColumn, button, shift, x, y, static_cast<HeaderHitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HotItemChanged(IListViewItem* pPreviousHotItem, IListViewItem* pNewHotItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HotItemChanged(pPreviousHotItem, pNewHotItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_HotItemChanging(IListViewItem* pPreviousHotItem, IListViewItem* pNewHotItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HotItemChanging(pPreviousHotItem, pNewHotItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_IncrementalSearching(BSTR currentSearchString, LONG* pItemToSelect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_IncrementalSearching(currentSearchString, pItemToSelect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_IncrementalSearchStringChanging(BSTR currentSearchString, SHORT keyCodeOfCharToBeAdded, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_IncrementalSearchStringChanging(currentSearchString, keyCodeOfCharToBeAdded, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InsertedColumn(IListViewColumn* pColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedColumn(pColumn);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InsertedGroup(IListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedGroup(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InsertedItem(IListViewItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedItem(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InsertingColumn(IVirtualListViewColumn* pColumn, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingColumn(pColumn, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InsertingGroup(IVirtualListViewGroup* pGroup, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingGroup(pGroup, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InsertingItem(IVirtualListViewItem* pListItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingItem(pListItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_InvokeVerbFromSubItemControl(IListViewItem* pListItem, BSTR verb)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InvokeVerbFromSubItemControl(pListItem, verb);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemActivate(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT shift, LONG x, LONG y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemActivate(pListItem, pListSubItem, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemAsynchronousDrawFailed(IListViewItem* pListItem, IListViewSubItem* pListSubItem, NMLVASYNCDRAWN* pNotificationDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemAsynchronousDrawFailed(pListItem, pListSubItem, pNotificationDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemBeginDrag(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginDrag(pListItem, pListSubItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemBeginRDrag(IListViewItem* pListItem, IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemBeginRDrag(pListItem, pListSubItem, button, shift, x, y, hitTestDetails);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemGetDisplayInfo(IListViewItem* pListItem, IListViewSubItem* pListSubItem, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pIndent, LONG* pGroupID, LPSAFEARRAY* ppTileViewColumns, LONG maxItemTextLength, BSTR* pItemText, LONG* pOverlayIndex, LONG* pStateImageIndex, ItemStateConstants* pItemStates, VARIANT_BOOL* pDontAskAgain)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetDisplayInfo(pListItem, pListSubItem, requestedInfo, pIconIndex, pIndent, pGroupID, ppTileViewColumns, maxItemTextLength, pItemText, pOverlayIndex, pStateImageIndex, pItemStates, pDontAskAgain);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemGetGroup(LONG itemIndex, LONG occurrenceIndex, LONG* pGroupIndex)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetGroup(itemIndex, occurrenceIndex, pGroupIndex);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemGetInfoTipText(IListViewItem* pListItem, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetInfoTipText(pListItem, maxInfoTipLength, pInfoTipText, pAbortToolTip);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemGetOccurrencesCount(LONG itemIndex, LONG* pOccurrencesCount)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemGetOccurrencesCount(itemIndex, pOccurrencesCount);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemMouseEnter(IListViewItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ItemMouseEnter(pListItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemMouseLeave(IListViewItem* pListItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ItemMouseLeave(pListItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemSelectionChanged(LVITEMINDEX& itemIndex)
{
	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		return Fire_ItemSelectionChanged(pLvwItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemSetText(IListViewItem* pListItem, BSTR itemText)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemSetText(pListItem, itemText);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemStateImageChanged(IListViewItem* pListItem, LONG previousStateImageIndex, LONG newStateImageIndex, StateImageChangeCausedByConstants causedBy)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemStateImageChanged(pListItem, previousStateImageIndex, newStateImageIndex, causedBy);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ItemStateImageChanging(IListViewItem* pListItem, LONG previousStateImageIndex, LONG* pNewStateImageIndex, StateImageChangeCausedByConstants causedBy, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ItemStateImageChanging(pListItem, previousStateImageIndex, pNewStateImageIndex, causedBy, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MapGroupWideToTotalItemIndex(LONG groupIndex, LONG groupWideItemIndex, LONG* pTotalItemIndex)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MapGroupWideToTotalItemIndex(groupIndex, groupWideItemIndex, pTotalItemIndex);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &mouseStatus.lastClickedItem, &subItemIndex);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(mouseStatus.lastClickedItem, subItemIndex, this, FALSE);
		}
		return Fire_MClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, mouseStatus.lastClickedItem.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);
	if(itemIndex.iItem != mouseStatus.lastClickedItem.iItem || itemIndex.iGroup != mouseStatus.lastClickedItem.iGroup) {
		itemIndex.iItem = -1;
	}
	mouseStatus.lastClickedItem.iItem = -1;
	mouseStatus.lastClickedItem.iGroup = 0;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}
		return Fire_MDblClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	if(!mouseStatus.hoveredControl) {
		if(!cachedSettings.hotTracking) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
	}
	mouseStatus.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		LVITEMINDEX itemIndex;
		int subItemIndex = -1;
		HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);

		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}

		return Fire_MouseDown(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!cachedSettings.hotTracking) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = *this;
		trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
		trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
		TrackMouseEvent(&trackingOptions);
	}

	mouseStatus.EnterControl();

	UINT h = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &h, &itemIndex, &subItemIndex);
	HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemIndex.iItem != -1);
	itemUnderMouse = itemIndex;
	subItemUnderMouse = subItemIndex;

	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pLvwItem && (subItemIndex > 0)) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		Fire_MouseEnter(pLvwItem, pLvwSubItem, button, shift, x, y, hitTestDetails);
	//}
	if(pLvwItem) {
		Raise_ItemMouseEnter(pLvwItem, button, shift, x, y, hitTestDetails);
		if(pLvwSubItem) {
			Raise_SubItemMouseEnter(pLvwSubItem, button, shift, x, y, hitTestDetails);
		}
	}
	return hr;
}

inline HRESULT ExplorerListView::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.hoveredControl) {
		mouseStatus.HoverControl();

		//if(m_nFreezeEvents == 0) {
			UINT hitTestDetails = 0;
			LVITEMINDEX itemIndex;
			int subItemIndex = -1;
			HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);

			CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
			CComPtr<IListViewSubItem> pLvwSubItem = NULL;
			if(pLvwItem && (subItemIndex > 0)) {
				pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
			}

			return Fire_MouseHover(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
		//}
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT h = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &h, &itemIndex, &subItemIndex);

	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
	if(pLvwItem) {
		HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
		if(subItemUnderMouse > 0) {
			CComPtr<IListViewSubItem> pLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
			Raise_SubItemMouseLeave(pLvwSubItem, button, shift, x, y, hitTestDetails);
		}
		Raise_ItemMouseLeave(pLvwItem, button, shift, x, y, hitTestDetails);
	}
	itemUnderMouse.iItem = -1;
	itemUnderMouse.iGroup = 0;
	subItemUnderMouse = -1;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}
		return Fire_MouseLeave(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(h, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT h = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &h, &itemIndex, &subItemIndex);

	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pLvwItem && (subItemIndex > 0)) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
	}

	// Do we move over another (sub-)item than before?
	if(itemIndex.iItem != itemUnderMouse.iItem || itemIndex.iGroup != itemUnderMouse.iGroup) {
		HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemUnderMouse.iItem != -1);
		CComPtr<IListViewItem> pPrevLvwItem = ClassFactory::InitListItem(itemUnderMouse, this);
		if(pPrevLvwItem) {
			if(subItemUnderMouse > 0) {
				CComPtr<IListViewSubItem> pPrevLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this, FALSE);
				Raise_SubItemMouseLeave(pPrevLvwSubItem, button, shift, x, y, hitTestDetails);
			}
			Raise_ItemMouseLeave(pPrevLvwItem, button, shift, x, y, hitTestDetails);
		}

		itemUnderMouse = itemIndex;
		subItemUnderMouse = subItemIndex;

		hitTestDetails = LVHTFlags2HitTestConstants(h, itemIndex.iItem != -1);
		if(pLvwItem) {
			Raise_ItemMouseEnter(pLvwItem, button, shift, x, y, hitTestDetails);
		}
		if(pLvwSubItem) {
			Raise_SubItemMouseEnter(pLvwSubItem, button, shift, x, y, hitTestDetails);
		}
	} else if(subItemIndex != subItemUnderMouse) {
		HitTestConstants hitTestDetails = LVHTFlags2HitTestConstants(h, itemIndex.iItem != -1);
		if(subItemUnderMouse > 0) {
			CComPtr<IListViewSubItem> pPrevLvwSubItem = ClassFactory::InitListSubItem(itemUnderMouse, subItemUnderMouse, this);
			Raise_SubItemMouseLeave(pPrevLvwSubItem, button, shift, x, y, hitTestDetails);
		}
		subItemUnderMouse = subItemIndex;
		if(pLvwSubItem) {
			Raise_SubItemMouseEnter(pLvwSubItem, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(h, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					/*if(!(properties.disabledEvents & deListClickEvents)) {
						Raise_Click(button, shift, x, y);
					}*/
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					/*if(!(properties.disabledEvents & deListClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}*/
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deListClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deListClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		LVITEMINDEX itemIndex;
		int subItemIndex = -1;
		HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);

		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}

		return Fire_MouseUp(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_MouseWheel(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, ScrollAxisConstants scrollAxis, SHORT wheelDelta)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	UINT h = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &h, &itemIndex, &subItemIndex);

	CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
	CComPtr<IListViewSubItem> pLvwSubItem = NULL;
	if(pLvwItem && (subItemIndex > 0)) {
		pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseWheel(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(h, itemIndex.iItem != -1), scrollAxis, wheelDelta);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGSCROLL);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	LVITEMINDEX dropTarget;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &dropTarget, NULL, TRUE);
	IListViewItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = FALSE;
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.dropTarget = ItemIndexToID(dropTarget.iItem);
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1);
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DROP, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
					dragDropStatus.drop_pDataObject = NULL;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<IListViewItem*>(eventDetails.pDropTarget);
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1));
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ExplorerListView::Raise_OLEDragEnter(BOOL fakedEnter, IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	LVITEMINDEX dropTarget;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &dropTarget, NULL, TRUE);
	IListViewItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CPoint mousePos(mousePosition.x, mousePosition.y);
		WTL::CRect noScrollZone;
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePos) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			if(containedSysHeader32.IsWindow() && containedSysHeader32.IsWindowVisible()) {
				WTL::CRect headerRectangle;
				containedSysHeader32.GetWindowRect(&headerRectangle);
				noScrollZone.top += headerRectangle.Height();
			}
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	if(!fakedEnter) {
		if(pData) {
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
	}
	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = FALSE;
			eventDetails.pDataObject = (fakedEnter ? NULL : reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p));
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.dropTarget = ItemIndexToID(dropTarget.iItem);
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1);
			eventDetails.autoHScrollVelocity = autoHScrollVelocity;
			eventDetails.autoVScrollVelocity = autoVScrollVelocity;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGENTER, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<IListViewItem*>(eventDetails.pDropTarget);
				autoHScrollVelocity = eventDetails.autoHScrollVelocity;
				autoVScrollVelocity = eventDetails.autoVScrollVelocity;
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject && !fakedEnter) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1), &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(pDropTarget) {
		// TODO: Retrieving just the index isn't enough!
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget.iItem = l;
		// we're using a raw pointer
		pDropTarget->Release();
	} else {
		dropTarget.iItem = -1;
		dropTarget.iGroup = 0;
	}

	dragDropStatus.lastDropTarget = dropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() >> 2 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OLEDragLeave(BOOL fakedLeave, BOOL* pCallDropTargetHelper)
{
	if(!fakedLeave) {
		KillTimer(timers.ID_DRAGSCROLL);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	UINT hitTestDetails = 0;
	LVITEMINDEX dropTarget;
	HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails, &dropTarget, NULL, TRUE);
	CComPtr<IListViewItem> pDropTarget = ClassFactory::InitListItem(dropTarget, this);

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = FALSE;
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget.p);
			eventDetails.dropTarget = ItemIndexToID(dropTarget.iItem);
			BUTTONANDSHIFT2OLEKEYSTATE(button, shift, eventDetails.keyState);
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = dragDropStatus.lastMousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1);
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGLEAVE, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1));
		}
	//}

	if(!fakedLeave) {
		dragDropStatus.pActiveDataObject = NULL;
		dragDropStatus.OLEDragLeaveOrDrop();
		Invalidate();
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition, BOOL* pCallDropTargetHelper)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	LVITEMINDEX dropTarget;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, &dropTarget, NULL, TRUE);
	IListViewItem* pDropTarget = NULL;
	ClassFactory::InitListItem(dropTarget, this, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	LONG autoHScrollVelocity = 0;
	LONG autoVScrollVelocity = 0;
	if(properties.dragScrollTimeBase != 0) {
		/* Use a 16 pixels wide border around the client area as the zone for auto-scrolling.
		   Are we within this zone? */
		WTL::CPoint mousePos(mousePosition.x, mousePosition.y);
		WTL::CRect noScrollZone;
		GetClientRect(&noScrollZone);
		BOOL isInScrollZone = (noScrollZone.PtInRect(mousePos) == TRUE);
		if(isInScrollZone) {
			// we're within the client area, so do further checks
			noScrollZone.DeflateRect(properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH, properties.DRAGSCROLLZONEWIDTH);
			if(containedSysHeader32.IsWindow() && containedSysHeader32.IsWindowVisible()) {
				WTL::CRect headerRectangle;
				containedSysHeader32.GetWindowRect(&headerRectangle);
				noScrollZone.top += headerRectangle.Height();
			}
			isInScrollZone = !noScrollZone.PtInRect(mousePos);
		}
		if(isInScrollZone) {
			// we're within the default scroll zone - propose some velocities
			if(mousePos.x < noScrollZone.left) {
				autoHScrollVelocity = -1;
			} else if(mousePos.x >= noScrollZone.right) {
				autoHScrollVelocity = 1;
			}
			if(mousePos.y < noScrollZone.top) {
				autoVScrollVelocity = -1;
			} else if(mousePos.y >= noScrollZone.bottom) {
				autoVScrollVelocity = 1;
			}
		}
	}

	HRESULT hr = S_OK;
	BOOL fireEvent = TRUE;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(shellBrowserInterface.pInternalMessageListener) {
			SHLVMDRAGDROPEVENTDATA eventDetails = {0};
			eventDetails.structSize = sizeof(SHLVMDRAGDROPEVENTDATA);
			eventDetails.headerIsTarget = FALSE;
			eventDetails.pDataObject = reinterpret_cast<LPUNKNOWN>(dragDropStatus.pActiveDataObject.p);
			eventDetails.effect = *pEffect;
			eventDetails.pDropTarget = reinterpret_cast<LPUNKNOWN>(pDropTarget);
			eventDetails.dropTarget = ItemIndexToID(dropTarget.iItem);
			eventDetails.keyState = keyState;
			eventDetails.draggingMouseButton = dragDropStatus.draggingMouseButton;
			eventDetails.cursorPosition = mousePosition;
			ClientToScreen(reinterpret_cast<LPPOINT>(&eventDetails.cursorPosition));
			eventDetails.hitTestDetails = LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1);
			eventDetails.autoHScrollVelocity = autoHScrollVelocity;
			eventDetails.autoVScrollVelocity = autoVScrollVelocity;
			eventDetails.pDropTargetHelper = dragDropStatus.pDropTargetHelper;

			hr = shellBrowserInterface.pInternalMessageListener->ProcessMessage(SHLVM_HANDLEDRAGEVENT, OLEDRAGEVENT_DRAGOVER, reinterpret_cast<LPARAM>(&eventDetails));
			if(hr == S_OK) {
				fireEvent = FALSE;
				if(pCallDropTargetHelper) {
					*pCallDropTargetHelper = FALSE;
				}

				*pEffect = eventDetails.effect;
				pDropTarget = reinterpret_cast<IListViewItem*>(eventDetails.pDropTarget);
				autoHScrollVelocity = eventDetails.autoHScrollVelocity;
				autoVScrollVelocity = eventDetails.autoVScrollVelocity;
			}
		}
	#endif

	//if(m_nFreezeEvents == 0) {
		if(fireEvent && dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, LVHTFlags2HitTestConstants(hitTestDetails, dropTarget.iItem != -1), &autoHScrollVelocity, &autoVScrollVelocity);
		}
	//}

	if(pDropTarget) {
		// TODO: Retrieving just the index isn't enough!
		LONG l = -1;
		pDropTarget->get_Index(&l);
		dropTarget.iItem = l;
		// we're using a raw pointer
		pDropTarget->Release();
	} else {
		dropTarget.iItem = -1;
		dropTarget.iGroup = 0;
	}

	dragDropStatus.lastDropTarget = dropTarget;

	if(properties.dragScrollTimeBase != 0) {
		dragDropStatus.autoScrolling.currentHScrollVelocity = autoHScrollVelocity;
		dragDropStatus.autoScrolling.currentVScrollVelocity = autoVScrollVelocity;

		LONG smallestInterval = max(abs(autoHScrollVelocity), abs(autoVScrollVelocity));
		if(smallestInterval) {
			smallestInterval = (properties.dragScrollTimeBase == -1 ? GetDoubleClickTime() >> 2 : properties.dragScrollTimeBase) / smallestInterval;
			if(smallestInterval == 0) {
				smallestInterval = 1;
			}
		}
		if(smallestInterval != dragDropStatus.autoScrolling.currentTimerInterval) {
			// reset the timer
			KillTimer(timers.ID_DRAGSCROLL);
			dragDropStatus.autoScrolling.currentTimerInterval = smallestInterval;
			if(smallestInterval != 0) {
				SetTimer(timers.ID_DRAGSCROLL, smallestInterval);
			}
		}
		if(smallestInterval) {
			/* Scroll immediately to avoid the theoretical situation where the timer interval is changed
			   faster than the timer fires so the control never is scrolled. */
			AutoScroll();
		}
	} else {
		KillTimer(timers.ID_DRAGSCROLL);
		dragDropStatus.autoScrolling.currentTimerInterval = 0;
	}

	return hr;
}

inline HRESULT ExplorerListView::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ExplorerListView::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ExplorerListView::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	if(dragDropStatus.headerIsSource) {
		return Raise_HeaderOLESetData(pData, formatID, index, dataOrViewAspect);
	} else /*if(m_nFreezeEvents == 0)*/ {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_OwnerDrawItem(IListViewItem* pListItem, OwnerDrawItemStateConstants itemState, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OwnerDrawItem(pListItem, itemState, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &mouseStatus.lastClickedItem, &subItemIndex);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(mouseStatus.lastClickedItem, subItemIndex, this, FALSE);
		}
		return Fire_RClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, mouseStatus.lastClickedItem.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);
	if(itemIndex.iItem != mouseStatus.lastClickedItem.iItem || itemIndex.iGroup != mouseStatus.lastClickedItem.iGroup) {
		itemIndex.iItem = -1;
	}
	mouseStatus.lastClickedItem.iItem = -1;
	mouseStatus.lastClickedItem.iGroup = 0;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}
		return Fire_RDblClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RecreatedControlWindow(HWND hWnd)
{
	// configure the listview control
	SendConfigurationMessages();
	if(containedSysHeader32.IsWindow()) {
		KillTimer(timers.ID_CREATED);
		Raise_CreatedHeaderControlWindow(containedSysHeader32);
	}

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	/* Comctl32.dll older than version 6.0 seems to simply toggle the internal flag instead of checking the
	   wParam value. So it's a bad idea to call SetRedraw(TRUE) when it already is set to TRUE. */
	if(properties.dontRedraw) {
		SetRedraw(!properties.dontRedraw);
	}
	SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RemovedColumn(IVirtualListViewColumn* pColumn)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedColumn(pColumn);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RemovedGroup(IVirtualListViewGroup* pGroup)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedGroup(pGroup);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RemovedItem(IVirtualListViewItem* pListItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedItem(pListItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RemovingColumn(IListViewColumn* pColumn, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingColumn(pColumn, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RemovingGroup(IListViewGroup* pGroup, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingGroup(pGroup, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RemovingItem(IListViewItem* pListItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingItem(pListItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RenamedItem(IListViewItem* pListItem, BSTR previousItemText, BSTR newItemText)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RenamedItem(pListItem, previousItemText, newItemText);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_RenamingItem(IListViewItem* pListItem, BSTR previousItemText, BSTR newItemText, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RenamingItem(pListItem, previousItemText, newItemText, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_ResizingColumn(IListViewColumn* pColumn, LONG* pNewColumnWidth, VARIANT_BOOL* pAbortResizing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizingColumn(pColumn, pNewColumnWidth, pAbortResizing);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_SelectedItemRange(IListViewItem* pFirstItem, IListViewItem* pLastItem)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectedItemRange(pFirstItem, pLastItem);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_SettingItemInfoTipText(IListViewItem* pListItem, BSTR* pInfoTipText, VARIANT_BOOL* pAbortInfoTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SettingItemInfoTipText(pListItem, pInfoTipText, pAbortInfoTip);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_StartingLabelEditing(IListViewItem* pListItem, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_StartingLabelEditing(pListItem, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_SubItemMouseEnter(IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_SubItemMouseEnter(pListSubItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_SubItemMouseLeave(IListViewSubItem* pListSubItem, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_SubItemMouseLeave(pListSubItem, button, shift, x, y, hitTestDetails);
	}
	return S_OK;
}

inline HRESULT ExplorerListView::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &mouseStatus.lastClickedItem, &subItemIndex);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(mouseStatus.lastClickedItem, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(mouseStatus.lastClickedItem, subItemIndex, this, FALSE);
		}
		return Fire_XClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, mouseStatus.lastClickedItem.iItem != -1));
	//}
	//return S_OK;
}

inline HRESULT ExplorerListView::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	LVITEMINDEX itemIndex;
	int subItemIndex = -1;
	HitTest(x, y, &hitTestDetails, &itemIndex, &subItemIndex);
	if(itemIndex.iItem != mouseStatus.lastClickedItem.iItem || itemIndex.iGroup != mouseStatus.lastClickedItem.iGroup) {
		itemIndex.iItem = -1;
	}
	mouseStatus.lastClickedItem.iItem = -1;
	mouseStatus.lastClickedItem.iGroup = 0;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, this);
		CComPtr<IListViewSubItem> pLvwSubItem = NULL;
		if(pLvwItem && (subItemIndex > 0)) {
			pLvwSubItem = ClassFactory::InitListSubItem(itemIndex, subItemIndex, this, FALSE);
		}
		return Fire_XDblClick(pLvwItem, pLvwSubItem, button, shift, x, y, LVHTFlags2HitTestConstants(hitTestDetails, itemIndex.iItem != -1));
	//}
	//return S_OK;
}


void ExplorerListView::RecreateControlWindow(void)
{
	static BOOL recursiveCall = FALSE;

	if(m_bInPlaceActive && !recursiveCall) {
		recursiveCall = TRUE;
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
		recursiveCall = FALSE;
	}
}

IMEModeConstants ExplorerListView::GetCurrentIMEContextMode(HWND hWnd)
{
	IMEModeConstants imeContextMode = imeNoControl;

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		imeContextMode = imeInherit;
	} else {
		HIMC hIMC = ImmGetContext(hWnd);
		if(hIMC) {
			if(ImmGetOpenStatus(hIMC)) {
				DWORD conversionMode = 0;
				DWORD sentenceMode = 0;
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				if(conversionMode & IME_CMODE_NATIVE) {
					if(conversionMode & IME_CMODE_KATAKANA) {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeKatakanaHalf];
						} else {
							imeContextMode = countryTable[imeAlphaFull];
						}
					} else {
						if(conversionMode & IME_CMODE_FULLSHAPE) {
							imeContextMode = countryTable[imeHiragana];
						} else {
							imeContextMode = countryTable[imeKatakana];
						}
					}
				} else {
					if(conversionMode & IME_CMODE_FULLSHAPE) {
						imeContextMode = countryTable[imeAlpha];
					} else {
						imeContextMode = countryTable[imeHangulFull];
					}
				}
			} else {
				imeContextMode = countryTable[imeDisable];
			}
			ImmReleaseContext(hWnd, hIMC);
		} else {
			imeContextMode = countryTable[imeOn];
		}
	}
	return imeContextMode;
}

void ExplorerListView::SetCurrentIMEContextMode(HWND hWnd, IMEModeConstants IMEMode)
{
	if((IMEMode == imeInherit) || (IMEMode == imeNoControl) || !::IsWindow(hWnd)) {
		return;
	}

	IMEModeConstants countryTable[10];
	IMEFlags.GetIMECountryTable(countryTable);
	if(countryTable[0] == -10) {
		return;
	}

	// update IME mode
	HIMC hIMC = ImmGetContext(hWnd);
	if(IMEMode == imeDisable) {
		// disable IME
		if(hIMC) {
			// close the IME
			if(ImmGetOpenStatus(hIMC)) {
				ImmSetOpenStatus(hIMC, FALSE);
			}
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
		// remove the control's association to the IME context
		HIMC h = ImmAssociateContext(hWnd, NULL);
		if(h) {
			IMEFlags.hDefaultIMC = h;
		}
		return;
	} else {
		// enable IME
		if(!hIMC) {
			if(!IMEFlags.hDefaultIMC) {
				// create an IME context
				hIMC = ImmCreateContext();
				if(hIMC) {
					// associate the control with the IME context
					ImmAssociateContext(hWnd, hIMC);
				}
			} else {
				// associate the control with the default IME context
				ImmAssociateContext(hWnd, IMEFlags.hDefaultIMC);
			}
		} else {
			// each ImmGetContext() needs a ImmReleaseContext()
			ImmReleaseContext(hWnd, hIMC);
			hIMC = NULL;
		}
	}

	hIMC = ImmGetContext(hWnd);
	if(hIMC) {
		DWORD conversionMode = 0;
		DWORD sentenceMode = 0;
		switch(IMEMode) {
			case imeOn:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				break;
			case imeOff:
				// close IME
				ImmSetOpenStatus(hIMC, FALSE);
				break;
			default:
				// open IME
				ImmSetOpenStatus(hIMC, TRUE);
				ImmGetConversionStatus(hIMC, &conversionMode, &sentenceMode);
				// switch conversion
				switch(IMEMode) {
					case imeHiragana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_KATAKANA;
						break;
					case imeKatakana:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeKatakanaHalf:
						conversionMode |= (IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
					case imeAlphaFull:
						conversionMode |= IME_CMODE_FULLSHAPE;
						conversionMode &= ~(IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeAlpha:
						conversionMode |= IME_CMODE_ALPHANUMERIC;
						conversionMode &= ~(IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE);
						break;
					case imeHangulFull:
						conversionMode |= (IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE);
						conversionMode &= ~IME_CMODE_ALPHANUMERIC;
						break;
					case imeHangul:
						conversionMode |= IME_CMODE_NATIVE;
						conversionMode &= ~IME_CMODE_FULLSHAPE;
						break;
				}
				ImmSetConversionStatus(hIMC, conversionMode, sentenceMode);
				break;
		}
		// each ImmGetContext() needs a ImmReleaseContext()
		ImmReleaseContext(hWnd, hIMC);
		hIMC = NULL;
	}
}

IMEModeConstants ExplorerListView::GetEffectiveIMEMode(void)
{
	IMEModeConstants IMEMode = properties.IMEMode;
	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

IMEModeConstants ExplorerListView::GetEffectiveIMEMode_Edit(void)
{
	IMEModeConstants IMEMode = properties.editIMEMode;
	if(IMEMode == imeInherit) {
		// retrieve the parent's IME mode
		IMEMode = GetEffectiveIMEMode();
	}

	if((IMEMode == imeInherit) && IsWindow()) {
		CWindow wnd = GetParent();
		while((IMEMode == imeInherit) && wnd.IsWindow()) {
			// retrieve the parent's IME mode
			IMEMode = GetCurrentIMEContextMode(wnd);
			wnd = wnd.GetParent();
		}
	}

	if(IMEMode == imeInherit) {
		// use imeNoControl as fallback
		IMEMode = imeNoControl;
	}
	return IMEMode;
}

DWORD ExplorerListView::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING | WS_EX_RIGHTSCROLLBAR;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD ExplorerListView::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	style |= LVS_SHAREIMAGELISTS;
	if(properties.allowLabelEditing) {
		style |= LVS_EDITLABELS;
	}
	if(properties.alwaysShowSelection) {
		style |= LVS_SHOWSELALWAYS;
	}
	switch(properties.autoArrangeItems) {
		case aaiIntelligent:
			if(IsComctl32Version610OrNewer()) {
				break;
			}
			// fall through
		case aaiNormal:
			style |= LVS_AUTOARRANGE;
			break;
	}
	if(!properties.clickableColumnHeaders) {
		style |= LVS_NOSORTHEADER;
	}
	switch(properties.columnHeaderVisibility) {
		case chvInvisible:
			style |= LVS_NOCOLUMNHEADER;
			break;
	}
	switch(properties.itemAlignment) {
		case iaTop:
			style |= LVS_ALIGNTOP;
			break;
		case iaLeft:
			style |= LVS_ALIGNLEFT;
			break;
		#ifdef ALLOWBOTTOMRIGHTALIGNMENT
			case iaBottom:
				style |= LVS_ALIGNBOTTOM;
				break;
			case iaRight:
				style |= LVS_ALIGNRIGHT;
				break;
		#endif
	}
	if(!properties.labelWrap) {
		style |= LVS_NOLABELWRAP;
	}
	if(!properties.multiSelect) {
		style |= LVS_SINGLESEL;
	}
	if(properties.ownerDrawn) {
		style |= LVS_OWNERDRAWFIXED;
	}
	switch(properties.scrollBars) {
		case sbNone:
			style |= LVS_NOSCROLL;
			break;
	}
	if(properties.virtualMode) {
		style |= LVS_OWNERDATA;
	}
	// force header control creation
	style |= LVS_REPORT;
	return style;
}

void ExplorerListView::SendConfigurationMessages(void)
{
	DWORD extendedStyle = 0;
	//BOOL comctl32580 = IsComctl32Version580OrNewer();
	BOOL comctl32600 = RunTimeHelper::IsCommCtrl6();
	BOOL comctl32610 = comctl32600 && IsComctl32Version610OrNewer();

	CComPtr<IListView_WIN7> pListView7 = NULL;
	CComPtr<IListView_WINVISTA> pListViewVista = NULL;
	SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WIN7), reinterpret_cast<LPARAM>(&pListView7));
	if(!pListView7) {
		SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_WINVISTA), reinterpret_cast<LPARAM>(&pListViewVista));
	}
	#ifdef INCLUDESUBITEMCONTROLINFRASTRUCTURE
		if(pListView7) {
			CComPtr<ISubItemCallback> pSubItemCallback = NULL;
			QueryInterface(IID_ISubItemCallback, reinterpret_cast<LPVOID*>(&pSubItemCallback));
			ATLVERIFY(SUCCEEDED(pListView7->SetSubItemCallback(pSubItemCallback)));
		} else if(pListViewVista) {
			CComPtr<ISubItemCallback> pSubItemCallback = NULL;
			QueryInterface(IID_ISubItemCallback, reinterpret_cast<LPVOID*>(&pSubItemCallback));
			ATLVERIFY(SUCCEEDED(pListViewVista->SetSubItemCallback(pSubItemCallback)));
		}
	#endif
	if(properties.virtualMode) {
		if(pListView7) {
			CComPtr<IOwnerDataCallback> pOwnerDataCallback = NULL;
			QueryInterface(IID_IOwnerDataCallback, reinterpret_cast<LPVOID*>(&pOwnerDataCallback));
			ATLVERIFY(SUCCEEDED(pListView7->SetOwnerDataCallback(pOwnerDataCallback)));
		} else if(pListViewVista) {
			ATLASSUME(pListViewVista);
			CComPtr<IOwnerDataCallback> pOwnerDataCallback = NULL;
			QueryInterface(IID_IOwnerDataCallback, reinterpret_cast<LPVOID*>(&pOwnerDataCallback));
			ATLVERIFY(SUCCEEDED(pListViewVista->SetOwnerDataCallback(pOwnerDataCallback)));
		}
	}

	if(comctl32600) {
		// now set the real view mode
		switch(properties.view) {
			case vIcons:
				SendMessage(LVM_SETVIEW, LV_VIEW_ICON, 0);
				break;
			case vSmallIcons:
				SendMessage(LVM_SETVIEW, LV_VIEW_SMALLICON, 0);
				break;
			case vList:
				SendMessage(LVM_SETVIEW, LV_VIEW_LIST, 0);
				break;
			case vDetails:
				SendMessage(LVM_SETVIEW, LV_VIEW_DETAILS, 0);
				break;
			case vExtendedTiles:
				if(comctl32610) {
					// the extended flag has to be set first
					LVTILEVIEWINFO tileViewInfo = {0};
					tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
					tileViewInfo.dwMask = LVTVIM_TILESIZE;
					tileViewInfo.dwFlags = LVTVIF_EXTENDED;
					SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
				}
				// fall through
			case vTiles:
				SendMessage(LVM_SETVIEW, LV_VIEW_TILE, 0);
				break;
		}

		if(comctl32610) {
			switch(properties.autoArrangeItems) {
				case aaiIntelligent:
					extendedStyle |= LVS_EX_AUTOAUTOARRANGE;
					break;
			}
		}
		if(properties.blendSelectionLasso) {
			extendedStyle |= LVS_EX_DOUBLEBUFFER;
		}
		if(properties.hideLabels) {
			extendedStyle |= LVS_EX_HIDELABELS;
		}
		if(properties.simpleSelect) {
			extendedStyle |= LVS_EX_SIMPLESELECT;
		}
		if(properties.singleRow) {
			extendedStyle |= LVS_EX_SINGLEROW;
		}
		if(properties.snapToGrid) {
			extendedStyle |= LVS_EX_SNAPTOGRID;
		}
	} else {
		// now set the real view mode
		switch(properties.view) {
			case vTiles:
			case vExtendedTiles:
			case vIcons:
				ModifyStyle(LVS_SMALLICON | LVS_LIST | LVS_REPORT, LVS_ICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				break;
			case vSmallIcons:
				ModifyStyle(LVS_ICON | LVS_LIST | LVS_REPORT, LVS_SMALLICON, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				break;
			case vList:
				ModifyStyle(LVS_ICON | LVS_SMALLICON | LVS_REPORT, LVS_LIST, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				break;
			case vDetails:
				ModifyStyle(LVS_ICON | LVS_SMALLICON | LVS_LIST, LVS_REPORT, SWP_DRAWFRAME | SWP_FRAMECHANGED);
				break;
		}
	}

	if(properties.allowHeaderDragDrop) {
		extendedStyle |= LVS_EX_HEADERDRAGDROP;
	}
	//if(comctl32580) {
		if(properties.borderSelect) {
			extendedStyle |= LVS_EX_BORDERSELECT;
		}
	//}
	if(properties.fullRowSelect != frsDisabled) {
		extendedStyle |= LVS_EX_FULLROWSELECT;
	}
	if(properties.gridLines) {
		extendedStyle |= LVS_EX_GRIDLINES;
	}
	if(properties.hotTracking) {
		extendedStyle |= LVS_EX_TRACKSELECT;
	}
	switch(properties.itemActivationMode) {
		case iamOneSingleClick:
			extendedStyle |= LVS_EX_ONECLICKACTIVATE;
			break;
		case iamTwoSingleClicks:
			extendedStyle |= LVS_EX_TWOCLICKACTIVATE;
			break;
	}
	if(properties.regional) {
		extendedStyle |= LVS_EX_REGIONAL;
	}
	switch(properties.scrollBars) {
		case sbFlat:
			extendedStyle |= LVS_EX_FLATSB;
			break;
	}
	if(properties.showStateImages) {
		extendedStyle |= LVS_EX_CHECKBOXES;
	}
	if(properties.showSubItemImages) {
		extendedStyle |= LVS_EX_SUBITEMIMAGES;
	}
	if(properties.toolTips & ttInfoTips) {
		extendedStyle |= LVS_EX_INFOTIP;
	}
	if((properties.toolTips & ttLabelTipsAlways)/* && comctl32580*/) {
		extendedStyle |= LVS_EX_LABELTIP;
	}
	if(properties.underlinedItems & uiHot) {
		extendedStyle |= LVS_EX_UNDERLINEHOT;
	}
	if(properties.underlinedItems & uiCold) {
		extendedStyle |= LVS_EX_UNDERLINECOLD;
	}
	if(properties.useWorkAreas) {
		extendedStyle |= LVS_EX_MULTIWORKAREAS;
	}
	if(comctl32610) {
		if(properties.autoSizeColumns) {
			extendedStyle |= LVS_EX_AUTOSIZECOLUMNS;
		}
		switch(properties.backgroundDrawMode) {
			case bdmByParent:
				extendedStyle |= LVS_EX_TRANSPARENTBKGND;
				break;
			case bdmByParentWithShadowText:
				extendedStyle |= LVS_EX_TRANSPARENTBKGND | LVS_EX_TRANSPARENTSHADOWTEXT;
				break;
		}
		if(properties.checkItemOnSelect) {
			extendedStyle |= LVS_EX_AUTOCHECKSELECT;
		}
		switch(properties.columnHeaderVisibility) {
			case chvVisibleInAllViews:
				extendedStyle |= LVS_EX_HEADERINALLVIEWS;
				break;
		}
		if(properties.drawImagesAsynchronously) {
			extendedStyle |= LVS_EX_DRAWIMAGEASYNC;
		}
		if(properties.justifyIconColumns) {
			extendedStyle |= LVS_EX_JUSTIFYCOLUMNS;
		}
		/*if(properties.showHeaderChevron) {
			extendedStyle |= LVS_EX_COLUMNOVERFLOW;
		}*/
		if(properties.useMinColumnWidths) {
			extendedStyle |= LVS_EX_COLUMNSNAPPOINTS;
		}
	}

	// call LVM_SETBKCOLOR before, because setting a back color clears LVS_EX_TRANSPARENTBKGND
	SendMessage(LVM_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
	SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE, 0, extendedStyle);

	SendMessage(LVM_SETTEXTCOLOR, 0, OLECOLOR2COLORREF(properties.foreColor));
	SendMessage(LVM_SETHOTLIGHTCOLOR, 0, (properties.hotForeColor == -1 ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.hotForeColor)));
	SendMessage(LVM_SETHOVERTIME, 0, properties.hotTrackingHoverTime);
	if(comctl32600) {
		SendMessage(LVM_ENABLEGROUPVIEW, properties.showGroups, 0);
		SendMessage(LVM_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
		SendMessage(LVM_SETOUTLINECOLOR, 0, OLECOLOR2COLORREF(properties.outlineColor));
		if(comctl32610) {
			CComPtr<IVisualProperties> pVisualProperties = NULL;
			if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IVisualProperties), reinterpret_cast<LPARAM>(&pVisualProperties))) {
				ATLASSUME(pVisualProperties);
				ATLVERIFY(SUCCEEDED(pVisualProperties->GetColor(VPCF_SORTCOLUMN, &properties.defaultSelectedColumnBackColor)));
				ATLVERIFY(SUCCEEDED(pVisualProperties->GetColor(VPCF_SUBTEXT, &properties.defaultTileViewSubItemForeColor)));
				if(properties.selectedColumnBackColor != -1) {
					ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SORTCOLUMN, OLECOLOR2COLORREF(properties.selectedColumnBackColor))));
				}
				if(properties.tileViewSubItemForeColor != -1) {
					ATLVERIFY(SUCCEEDED(pVisualProperties->SetColor(VPCF_SUBTEXT, OLECOLOR2COLORREF(properties.tileViewSubItemForeColor))));
				}
			}
		}

		LVGROUPMETRICS groupMetrics = {0};
		groupMetrics.cbSize = sizeof(LVGROUPMETRICS);
		if(comctl32610) {
			groupMetrics.mask |= LVGMF_BORDERSIZE;
			groupMetrics.Bottom = properties.groupMarginBottom;
			groupMetrics.Left = properties.groupMarginLeft;
			groupMetrics.Right = properties.groupMarginRight;
			groupMetrics.Top = properties.groupMarginTop;
		}
		groupMetrics.mask |= LVGMF_TEXTCOLOR;
		groupMetrics.crFooter = OLECOLOR2COLORREF(properties.groupFooterForeColor);
		groupMetrics.crHeader = OLECOLOR2COLORREF(properties.groupHeaderForeColor);
		SendMessage(LVM_SETGROUPMETRICS, 0, reinterpret_cast<LPARAM>(&groupMetrics));

		LVTILEVIEWINFO tileViewInfo = {0};
		tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
		tileViewInfo.dwMask = LVTVIM_COLUMNS | LVTVIM_LABELMARGIN;
		tileViewInfo.cLines = properties.tileViewItemLines;
		tileViewInfo.rcLabelMargin.bottom = properties.tileViewLabelMarginBottom;
		tileViewInfo.rcLabelMargin.left = properties.tileViewLabelMarginLeft;
		tileViewInfo.rcLabelMargin.right = properties.tileViewLabelMarginRight;
		tileViewInfo.rcLabelMargin.top = properties.tileViewLabelMarginTop;
		if(properties.tileViewTileHeight != -1) {
			tileViewInfo.sizeTile.cy = properties.tileViewTileHeight;
			tileViewInfo.dwFlags |= LVTVIF_FIXEDHEIGHT;
			tileViewInfo.dwMask |= LVTVIM_TILESIZE;
		}
		if(properties.tileViewTileWidth != -1) {
			tileViewInfo.sizeTile.cx = properties.tileViewTileWidth;
			tileViewInfo.dwFlags |= LVTVIF_FIXEDWIDTH;
			tileViewInfo.dwMask |= LVTVIM_TILESIZE;
		}
		if(comctl32610) {
			if(properties.view == vExtendedTiles) {
				tileViewInfo.dwFlags |= LVTVIF_EXTENDED;
			}
		}
		SendMessage(LVM_SETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
	}
	SendMessage(LVM_SETTEXTBKCOLOR, 0, (properties.textBackColor == CLR_NONE ? CLR_NONE : OLECOLOR2COLORREF(properties.textBackColor)));

	// set the hot cursor
	if(properties.hotMousePointer == mpCustom) {
		if(properties.hotMouseIcon.pPictureDisp) {
			CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.hotMouseIcon.pPictureDisp);
			if(pPicture) {
				OLE_HANDLE hCursor = NULL;
				pPicture->get_Handle(&hCursor);
				properties.hotMouseIcon.dontGetPictureObject = TRUE;
				SendMessage(LVM_SETHOTCURSOR, 0, reinterpret_cast<LPARAM>(LongToHandle(hCursor)));
				properties.hotMouseIcon.dontGetPictureObject = FALSE;
			}
		}
	} else if(properties.hotMousePointer != mpDefault) {
		HCURSOR hCursor = NULL;
		hCursor = MousePointerConst2hCursor(properties.hotMousePointer);
		properties.hotMouseIcon.dontGetPictureObject = TRUE;
		SendMessage(LVM_SETHOTCURSOR, 0, reinterpret_cast<LPARAM>(hCursor));
		properties.hotMouseIcon.dontGetPictureObject = FALSE;
	}

	// set the background image
	LVBKIMAGE bkImage = {0};
	BSTR tmp = NULL;
	if(properties.bkImage.vt == VT_BSTR) {
		tmp = properties.bkImage.bstrVal;
	}
	CW2T converter(tmp);
	switch(properties.bkImage.vt) {
		case VT_BSTR:
			bkImage.pszImage = converter;
			bkImage.cchImageMax = lstrlen(bkImage.pszImage);
			break;
		case VT_DISPATCH:
			if(properties.bkImage.pdispVal) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.bkImage.pdispVal);
				if(pPicture) {
					SHORT pictureType = PICTYPE_NONE;
					pPicture->get_Type(&pictureType);
					if(pictureType == PICTYPE_BITMAP) {
						OLE_HANDLE h = NULL;
						pPicture->get_Handle(&h);

						// SysListView32 always needs its own bitmap
						bkImage.hbm = CopyBitmap(static_cast<HBITMAP>(LongToHandle(h)));
					}
				}
			}
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			if(SUCCEEDED(VariantChangeType(&v, &properties.bkImage, 0, VT_I4))) {
				bkImage.hbm = static_cast<HBITMAP>(LongToHandle(v.lVal));
			}
			break;
	}
	bkImage.xOffsetPercent = properties.bkImagePositionX;
	bkImage.yOffsetPercent = properties.bkImagePositionY;
	if(properties.absoluteBkImagePosition) {
		bkImage.ulFlags |= LVBKIF_FLAG_TILEOFFSET;
	}

	LRESULT lr = TRUE;
	if(bkImage.cchImageMax) {
		bkImage.ulFlags |= LVBKIF_SOURCE_URL;
		switch(properties.bkImageStyle) {
			case bisWatermark:
				// not supported - fall through
			case bisNormal:
				bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
				break;
			case bisTiled:
				bkImage.ulFlags |= LVBKIF_STYLE_TILE;
				break;
		}

		// NOTE: We had wParam set to TRUE here. Why?
		lr = SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
	} else if(bkImage.hbm) {
		if(comctl32600) {
			bkImage.ulFlags |= LVBKIF_SOURCE_HBITMAP;
			switch(properties.bkImageStyle) {
				case bisNormal:
					bkImage.ulFlags |= LVBKIF_STYLE_NORMAL;
					break;
				case bisTiled:
					bkImage.ulFlags |= LVBKIF_STYLE_TILE;
					break;
				case bisWatermark:
					// yes, we've to set LVBKIF_SOURCE_NONE - the MS employee must have been stoned...
					bkImage.ulFlags = LVBKIF_SOURCE_NONE | LVBKIF_STYLE_NORMAL | LVBKIF_TYPE_WATERMARK;
					break;
			}

			// NOTE: We had wParam set to TRUE here. Why?
			lr = SendMessage(LVM_SETBKIMAGE, 0, reinterpret_cast<LPARAM>(&bkImage));
		} else {
			lr = 0;
		}
	}
	if(!lr && !IsInDesignMode()) {
		// something went wrong while setting the background image - synchronize the BkImage property
		if(bkImage.hbm && (GetObjectType(bkImage.hbm) == OBJ_BITMAP)) {
			DeleteObject(bkImage.hbm);
		}
		VariantClear(&properties.bkImage);
	}

	if(pListView7) {
		ATLVERIFY(SUCCEEDED(pListView7->SetGroupSubsetCount(properties.minItemRowsVisibleInGroups)));
		if(properties.fullRowSelect == frsExtendedMode) {
			ATLVERIFY(SUCCEEDED(pListView7->SetSelectionFlags(0x01, 0x01)));
		}
		FireViewChange();
	} else if(pListViewVista) {
		ATLVERIFY(SUCCEEDED(pListViewVista->SetGroupSubsetCount(properties.minItemRowsVisibleInGroups)));
		if(properties.fullRowSelect == frsExtendedMode) {
			ATLVERIFY(SUCCEEDED(pListViewVista->SetSelectionFlags(0x01, 0x01)));
		}
		FireViewChange();
	}

	ApplyFont();

	SendMessage(LVM_SETCALLBACKMASK, properties.callBackMask, 0);
	if(properties.virtualMode) {
		SendMessage(LVM_SETITEMCOUNT, properties.virtualItemCount, LVSICF_NOINVALIDATEALL);
	}

	if(IsInDesignMode() && !(GetStyle() & LVS_OWNERDATA)) {
		LVCOLUMN column = {0};
		column.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
		column.cx = 130;
		column.fmt = LVCFMT_LEFT;
		column.pszText = TEXT("Dummy Column 1");
		column.cchTextMax = lstrlen(column.pszText);
		SendMessage(LVM_INSERTCOLUMN, 0, reinterpret_cast<LPARAM>(&column));
		column.fmt = LVCFMT_RIGHT;
		column.pszText = TEXT("Dummy Column 2");
		column.cchTextMax = lstrlen(column.pszText);
		SendMessage(LVM_INSERTCOLUMN, 1, reinterpret_cast<LPARAM>(&column));

		if(comctl32600) {
			if(properties.showGroups) {
				LVGROUP group = {0};
				group.cbSize = RunTimeHelper::SizeOf_LVGROUP();
				group.mask = LVGF_ALIGN | LVGF_GROUPID | LVGF_HEADER;
				group.uAlign = LVGA_HEADER_LEFT;
				group.iGroupId = 1;
				group.pszHeader = L"Dummy Group 1";
				group.cchHeader = lstrlenW(group.pszHeader);
				SendMessage(LVM_INSERTGROUP, 0, reinterpret_cast<LPARAM>(&group));
				group.uAlign = LVGA_HEADER_RIGHT;
				group.iGroupId = 2;
				group.pszHeader = L"Dummy Group 2";
				group.cchHeader = lstrlenW(group.pszHeader);
				SendMessage(LVM_INSERTGROUP, 1, reinterpret_cast<LPARAM>(&group));
			}
		}

		UINT columns[1] = {1};
		LVITEM insertionData = {0};
		insertionData.mask = LVIF_TEXT;
		if(comctl32600) {
			if(properties.showGroups) {
				insertionData.mask |= LVIF_GROUPID;
				insertionData.iGroupId = 1;
			}
			insertionData.mask |= LVIF_COLUMNS;
			insertionData.cColumns = 1;
			insertionData.puColumns = columns;
		}
		insertionData.iItem = 0;
		insertionData.pszText = TEXT("Dummy Item 1");
		insertionData.cchTextMax = lstrlen(insertionData.pszText);
		SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
		insertionData.iItem = 1;
		insertionData.pszText = TEXT("Dummy Item 2");
		insertionData.cchTextMax = lstrlen(insertionData.pszText);
		SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
		insertionData.iItem = 2;
		insertionData.pszText = TEXT("Dummy Item 3");
		insertionData.cchTextMax = lstrlen(insertionData.pszText);
		SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
		if(comctl32600 && properties.showGroups) {
			insertionData.iGroupId = 2;
		}
		insertionData.iItem = 3;
		insertionData.pszText = TEXT("Dummy Item 4");
		insertionData.cchTextMax = lstrlen(insertionData.pszText);
		SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
		insertionData.iItem = 4;
		insertionData.pszText = TEXT("Dummy Item 5");
		insertionData.cchTextMax = lstrlen(insertionData.pszText);
		SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
		insertionData.iItem = 5;
		insertionData.pszText = TEXT("Dummy Item 6");
		insertionData.cchTextMax = lstrlen(insertionData.pszText);
		SendMessage(LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData));

		LVITEM subItem = {0};
		subItem.iSubItem = 1;
		subItem.pszText = TEXT("Dummy Sub Item 1");
		subItem.cchTextMax = lstrlen(subItem.pszText);
		SendMessage(LVM_SETITEMTEXT, 0, reinterpret_cast<LPARAM>(&subItem));
		subItem.pszText = TEXT("Dummy Sub Item 2");
		subItem.cchTextMax = lstrlen(subItem.pszText);
		SendMessage(LVM_SETITEMTEXT, 1, reinterpret_cast<LPARAM>(&subItem));
		subItem.pszText = TEXT("Dummy Sub Item 3");
		subItem.cchTextMax = lstrlen(subItem.pszText);
		SendMessage(LVM_SETITEMTEXT, 2, reinterpret_cast<LPARAM>(&subItem));
		subItem.pszText = TEXT("Dummy Sub Item 4");
		subItem.cchTextMax = lstrlen(subItem.pszText);
		SendMessage(LVM_SETITEMTEXT, 3, reinterpret_cast<LPARAM>(&subItem));
		subItem.pszText = TEXT("Dummy Sub Item 5");
		subItem.cchTextMax = lstrlen(subItem.pszText);
		SendMessage(LVM_SETITEMTEXT, 4, reinterpret_cast<LPARAM>(&subItem));
		subItem.pszText = TEXT("Dummy Sub Item 6");
		subItem.cchTextMax = lstrlen(subItem.pszText);
		SendMessage(LVM_SETITEMTEXT, 5, reinterpret_cast<LPARAM>(&subItem));
	}

	if(comctl32610 && IsInDesignMode()) {
		CComPtr<IListViewFooter> pListViewFooter = NULL;
		if(SendMessage(LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListViewFooter), reinterpret_cast<LPARAM>(&pListViewFooter))) {
			ATLASSUME(pListViewFooter);
			pListViewFooter->SetIntroText(L"Dummy Footer Items");
			CComPtr<IListViewFooterCallback> pListViewFooterCallback = NULL;
			QueryInterface(IID_IListViewFooterCallback, reinterpret_cast<LPVOID*>(&pListViewFooterCallback));
			pListViewFooter->Show(pListViewFooterCallback);
			pListViewFooter->InsertButton(0, L"Button 1", NULL, 0, 0);
			pListViewFooter->InsertButton(1, L"Button 2", NULL, 0, 0);
		}
	}
}

HCURSOR ExplorerListView::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

int ExplorerListView::IDToColumnIndex(LONG ID)
{
	#ifdef USE_STL
		std::unordered_map< LONG, int, ItemIndexHasher<LONG> >::iterator iter = columnIndexes.find(ID);
		if(iter != columnIndexes.end()) {
			return iter->second;
		}
	#else
		const CAtlMap<LONG, int>::CPair* p = columnIndexes.Lookup(ID);
		if(p) {
			return p->m_value;
		}
	#endif
	return -1;
}

LONG ExplorerListView::ColumnIndexToID(int columnIndex)
{
	ATLASSERT(containedSysHeader32.IsWindow());

	HDITEM column = {0};
	column.mask = HDI_LPARAM | HDI_NOINTERCEPTION;
	if(containedSysHeader32.SendMessage(HDM_GETITEM, columnIndex, reinterpret_cast<LPARAM>(&column))) {
		return static_cast<LONG>(column.lParam);
	}
	return -1;
}

int ExplorerListView::IDToItemIndex(LONG ID)
{
	if(RunTimeHelper::IsCommCtrl6()) {
		ATLASSERT(IsWindow());

		return static_cast<int>(SendMessage(LVM_MAPIDTOINDEX, ID, 0));
	} else {
		#ifdef USE_STL
			std::vector<LONG>::iterator iter = std::find(itemIDs.begin(), itemIDs.end(), ID);
			if(iter != itemIDs.end()) {
				return std::distance(itemIDs.begin(), iter);
			}
		#else
			for(size_t i = 0; i < itemIDs.GetCount(); ++i) {
				if(itemIDs[i] == ID) {
					return i;
				}
			}
		#endif
	}
	return -1;
}

LONG ExplorerListView::ItemIndexToID(int itemIndex)
{
	ATLASSERT(IsWindow());

	if(RunTimeHelper::IsCommCtrl6()) {
		return static_cast<LONG>(SendMessage(LVM_MAPINDEXTOID, itemIndex, 0));
	} else {
		LVITEM item = {0};
		item.iItem = itemIndex;
		item.mask = LVIF_PARAM | LVIF_NOINTERCEPTION;
		if(SendMessage(LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			return static_cast<LONG>(item.lParam);
		}
	}
	return -1;
}

int ExplorerListView::GroupIDToPositionIndex(int groupID)
{
	#ifdef USE_STL
		std::vector<int>::iterator iter = std::find(groups.begin(), groups.end(), groupID);
		if(iter != groups.end()) {
			return static_cast<int>(std::distance(groups.begin(), iter));
		}
	#else
		for(size_t i = 0; i < groups.GetCount(); ++i) {
			if(groups[i] == groupID) {
				return static_cast<int>(i);
			}
		}
	#endif
	return -1;
}

int ExplorerListView::PositionIndexToGroupID(int positionIndex)
{
	#ifdef USE_STL
		if((positionIndex >= 0) && (positionIndex < static_cast<int>(groups.size()))) {
	#else
		if((positionIndex >= 0) && (positionIndex < static_cast<int>(groups.GetCount()))) {
	#endif
		return groups[positionIndex];
	}
	return -1;
}

UINT ExplorerListView::HitTestConstants2LVHTFlags(HitTestConstants bitField)
{
	UINT translated = static_cast<UINT>(bitField);
	// re-set all flags that are not just mapped (htItem is handled through htItemStateImage)
	translated &= ~(0x0040/*htItemStateImage*/ | 0x0100/*htAbove*/ | 0x0200/*htBelow*/ | 0x0400/*htToRight*/ | 0x0800/*htToLeft*/);
	if(bitField & htItemStateImage) {
		translated |= LVHT_ONITEMSTATEICON;
	}
	if(bitField & htAbove) {
		translated |= LVHT_ABOVE;
	}
	if(bitField & htBelow) {
		translated |= LVHT_BELOW;
	}
	if(bitField & htToRight) {
		translated |= LVHT_TORIGHT;
	}
	if(bitField & htToLeft) {
		translated |= LVHT_TOLEFT;
	}

	return translated;
}

HitTestConstants ExplorerListView::LVHTFlags2HitTestConstants(UINT bitField, BOOL itemWasHit)
{
	HitTestConstants translated = static_cast<HitTestConstants>(bitField & ~(LVHT_ONITEMSTATEICON | LVHT_ABOVE | LVHT_BELOW | LVHT_TORIGHT | LVHT_TOLEFT));
	if(bitField & 0x0008/*LVHT_ONITEMSTATEICON/LVHT_ABOVE*/) {
		translated = static_cast<HitTestConstants>(translated | (itemWasHit ? htItemStateImage : htAbove));
	}
	if(bitField & LVHT_BELOW) {
		translated = static_cast<HitTestConstants>(translated | htBelow);
	}
	if(bitField & LVHT_TORIGHT) {
		translated = static_cast<HitTestConstants>(translated | htToRight);
	}
	if(bitField & LVHT_TOLEFT) {
		translated = static_cast<HitTestConstants>(translated | htToLeft);
	}

	return translated;
}

LONG ExplorerListView::GetNewItemID(void)
{
	static LONG lastID = -1;

	return ++lastID;
}

void ExplorerListView::RegisterItemContainer(IItemContainer* pContainer)
{
	ATLASSUME(pContainer);
	#ifdef DEBUG
		#ifdef USE_STL
			std::unordered_map<DWORD, IItemContainer*>::iterator iter = itemContainers.find(pContainer->GetID());
			ATLASSERT(iter == itemContainers.end());
		#else
			CAtlMap<DWORD, IItemContainer*>::CPair* pEntry = itemContainers.Lookup(pContainer->GetID());
			ATLASSERT(pEntry == NULL);
		#endif
	#endif
	itemContainers[pContainer->GetID()] = pContainer;
}

void ExplorerListView::DeregisterItemContainer(DWORD containerID)
{
	#ifdef USE_STL
		std::unordered_map<DWORD, IItemContainer*>::iterator iter = itemContainers.find(containerID);
		ATLASSERT(iter != itemContainers.end());
		if(iter != itemContainers.end()) {
			itemContainers.erase(iter);
		}
	#else
		itemContainers.RemoveKey(containerID);
	#endif
}

void ExplorerListView::RemoveItemFromItemContainers(LONG itemIdentifier)
{
	#ifdef USE_STL
		for(std::unordered_map<DWORD, IItemContainer*>::const_iterator iter = itemContainers.begin(); iter != itemContainers.end(); ++iter) {
			iter->second->RemovedItem(itemIdentifier);
		}
	#else
		POSITION p = itemContainers.GetStartPosition();
		while(p) {
			itemContainers.GetValueAt(p)->RemovedItem(itemIdentifier);
			itemContainers.GetNextValue(p);
		}
	#endif
}

LONG ExplorerListView::GetNewColumnID(void)
{
	static LONG lastID = -1;

	return ++lastID;
}

LCID ExplorerListView::GetColumnLocale(int columnIndex)
{
	#ifdef USE_STL
		std::list<ColumnData>::iterator iter = columnParams.begin();
		if(iter != columnParams.end()) {
			std::advance(iter, columnIndex);
			if(iter != columnParams.end()) {
				return iter->localeID;
			}
		}
	#else
		POSITION p = columnParams.FindIndex(columnIndex);
		if(p) {
			return columnParams.GetAt(p).localeID;
		}
	#endif
	return static_cast<LCID>(-1);
}

void ExplorerListView::SetColumnLocale(int columnIndex, LCID localeID)
{
	#ifdef USE_STL
		std::list<ColumnData>::iterator iter = columnParams.begin();
		if(iter != columnParams.end()) {
			std::advance(iter, columnIndex);
			if(iter != columnParams.end()) {
				iter->localeID = localeID;
			}
		}
	#else
		POSITION p = columnParams.FindIndex(columnIndex);
		if(p) {
			columnParams.GetAt(p).localeID = localeID;
		}
	#endif
}

DWORD ExplorerListView::GetColumnTextParsingFlags(int columnIndex, TextParsingFunctionConstants parsingFunction)
{
	#ifdef USE_STL
		std::list<ColumnData>::iterator iter = columnParams.begin();
		if(iter != columnParams.end()) {
			std::advance(iter, columnIndex);
			if(iter != columnParams.end()) {
				switch(parsingFunction) {
					case tpfCompareString:
						return iter->flagsForCompareString;
						break;
					case tpfVarFromStr:
						return iter->flagsForVarFromStr;
						break;
				}
			}
		}
	#else
		POSITION p = columnParams.FindIndex(columnIndex);
		if(p) {
			switch(parsingFunction) {
				case tpfCompareString:
					return columnParams.GetAt(p).flagsForCompareString;
					break;
				case tpfVarFromStr:
					return columnParams.GetAt(p).flagsForVarFromStr;
					break;
			}
		}
	#endif
	return 0;
}

void ExplorerListView::SetColumnTextParsingFlags(int columnIndex, TextParsingFunctionConstants parsingFunction, DWORD options)
{
	#ifdef USE_STL
		std::list<ColumnData>::iterator iter = columnParams.begin();
		if(iter != columnParams.end()) {
			std::advance(iter, columnIndex);
			if(iter != columnParams.end()) {
				switch(parsingFunction) {
					case tpfCompareString:
						iter->flagsForCompareString = options;
						break;
					case tpfVarFromStr:
						iter->flagsForVarFromStr = options;
						break;
				}
			}
		}
	#else
		POSITION p = columnParams.FindIndex(columnIndex);
		if(p) {
			switch(parsingFunction) {
				case tpfCompareString:
					columnParams.GetAt(p).flagsForCompareString = options;
					break;
				case tpfVarFromStr:
					columnParams.GetAt(p).flagsForVarFromStr = options;
					break;
			}
		}
	#endif
}

int ExplorerListView::HeaderHitTest(LONG x, LONG y, UINT* pFlags)
{
	ATLASSERT(containedSysHeader32.IsWindow());

	UINT flags = 0;
	if(pFlags) {
		flags = *pFlags;
	}
	HDHITTESTINFO hitTestInfo = {{x, y}, flags, -1};
	int column = static_cast<int>(containedSysHeader32.SendMessage(HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo)));
	flags = hitTestInfo.flags;
	if(pFlags) {
		*pFlags = flags;
	}
	return column;
}

void ExplorerListView::HitTest(LONG x, LONG y, UINT* pFlags, LVITEMINDEX* pHitItem, PINT pHitSubItem, BOOL ignoreBoundingBoxDefinition/* = FALSE*/, BOOL wantExtendedFlags/* = TRUE*/, BOOL wantItemIndex/* = TRUE*/)
{
	ATLASSERT(IsWindow());
	ATLASSERT(wantExtendedFlags || wantItemIndex);

	BOOL comctl32Version610 = IsComctl32Version610OrNewer();
	wantExtendedFlags = wantExtendedFlags && comctl32Version610;
	wantItemIndex = wantItemIndex || !comctl32Version610;

	LVITEMINDEX dummy;
	if(!pHitItem) {
		pHitItem = &dummy;
	}

	UINT flags = 0;
	UINT normalFlags = 0;
	if(pFlags) {
		flags = *pFlags;
	}
	LVHITTESTINFO hitTestInfo = {{x, y}, flags, 0, 0, 0};
	pHitItem->iItem = -1;
	pHitItem->iGroup = 0;
	if(pHitSubItem) {
		pHitItem->iItem = static_cast<int>(SendMessage(LVM_SUBITEMHITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo)));
		/* NOTE: Starting with Vista, iSubItem is set to the column under the mouse if the control is empty.
		         This might depend on whether an empty markup text is specified and whether this empty markup
		         text contains links. */
		if(pHitItem->iItem == -1) {
			hitTestInfo.iSubItem = 0;
		}
		pHitItem->iGroup = hitTestInfo.iGroup;
		*pHitSubItem = hitTestInfo.iSubItem;
		normalFlags = hitTestInfo.flags;

		if(wantExtendedFlags) {
			hitTestInfo.flags = flags;
			if(wantItemIndex) {
				SendMessage(LVM_SUBITEMHITTEST, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&hitTestInfo));
			} else {
				// the caller wants whatever LVM_SUBITEMHITTEST returns
				pHitItem->iItem = static_cast<int>(SendMessage(LVM_SUBITEMHITTEST, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&hitTestInfo)));
				pHitItem->iGroup = hitTestInfo.iGroup;
				*pHitSubItem = hitTestInfo.iSubItem;
			}
			hitTestInfo.flags |= normalFlags;
		}
	} else {
		pHitItem->iItem = static_cast<int>(SendMessage(LVM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo)));
		pHitItem->iGroup = hitTestInfo.iGroup;
		normalFlags = hitTestInfo.flags;

		if(wantExtendedFlags) {
			hitTestInfo.flags = flags;
			if(wantItemIndex) {
				SendMessage(LVM_HITTEST, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&hitTestInfo));
			} else {
				// the caller wants whatever LVM_HITTEST returns
				pHitItem->iItem = static_cast<int>(SendMessage(LVM_HITTEST, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&hitTestInfo)));
				pHitItem->iGroup = hitTestInfo.iGroup;
			}
			hitTestInfo.flags |= normalFlags;
		}
	}
	if(comctl32Version610) {
		// HACK: On Vista RTM, if the point is outside the control, we get LVHT_NOWHERE.
		RECT clientRectangle = {0};
		GetClientRect(&clientRectangle);
		if(x < clientRectangle.left) {
			hitTestInfo.flags &= ~LVHT_NOWHERE;
			hitTestInfo.flags |= LVHT_TOLEFT;
		} else if(x >= clientRectangle.right) {
			hitTestInfo.flags &= ~LVHT_NOWHERE;
			hitTestInfo.flags |= LVHT_TORIGHT;
		}
		if(y < clientRectangle.top) {
			hitTestInfo.flags &= ~LVHT_NOWHERE;
			hitTestInfo.flags |= LVHT_ABOVE;
		} else if(y >= clientRectangle.bottom) {
			hitTestInfo.flags &= ~LVHT_NOWHERE;
			hitTestInfo.flags |= LVHT_BELOW;
		}
	}
	flags = hitTestInfo.flags;
	if(pFlags) {
		*pFlags = flags;
	}
	if(!ignoreBoundingBoxDefinition && wantItemIndex && ((HitTestConstants2LVHTFlags(static_cast<HitTestConstants>(properties.itemBoundingBoxDefinition)) & normalFlags) != normalFlags)) {
		pHitItem->iItem = -1;
		pHitItem->iGroup = 0;
	}
}

BOOL ExplorerListView::IsInDesignMode(void)
{
	BOOL b = TRUE;
	this->GetAmbientUserMode(b);
	return !b;
}

void ExplorerListView::AutoScroll(void)
{
	LONG realScrollTimeBase = properties.dragScrollTimeBase;
	if(realScrollTimeBase == -1) {
		realScrollTimeBase = GetDoubleClickTime() >> 2;
	}

	if((dragDropStatus.autoScrolling.currentHScrollVelocity != 0) && ((GetStyle() & WS_HSCROLL) == WS_HSCROLL)) {
		if(dragDropStatus.autoScrolling.currentHScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll to the left?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Left) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll left
					dragDropStatus.autoScrolling.lastScroll_Left = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINELEFT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll to the right?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Right) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentHScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE;
				GetScrollInfo(SB_HORZ, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll right
					dragDropStatus.autoScrolling.lastScroll_Right = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_HSCROLL, SB_LINERIGHT, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}

	if((dragDropStatus.autoScrolling.currentVScrollVelocity != 0) && ((GetStyle() & WS_VSCROLL) == WS_VSCROLL)) {
		if(dragDropStatus.autoScrolling.currentVScrollVelocity < 0) {
			// Have we been waiting long enough since the last scroll upwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Up) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPos > scrollInfo.nMin) {
					// scroll up
					dragDropStatus.autoScrolling.lastScroll_Up = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEUP, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		} else {
			// Have we been waiting long enough since the last scroll downwards?
			if((GetTickCount() - dragDropStatus.autoScrolling.lastScroll_Down) >= static_cast<ULONG>(realScrollTimeBase / abs(dragDropStatus.autoScrolling.currentVScrollVelocity))) {
				SCROLLINFO scrollInfo = {0};
				scrollInfo.cbSize = sizeof(SCROLLINFO);
				scrollInfo.fMask = SIF_POS | SIF_RANGE | SIF_PAGE;
				GetScrollInfo(SB_VERT, &scrollInfo);
				if(scrollInfo.nPage) {
					scrollInfo.nMax -= scrollInfo.nPage - 1;
				}
				if(scrollInfo.nPos < scrollInfo.nMax) {
					// scroll down
					dragDropStatus.autoScrolling.lastScroll_Down = GetTickCount();
					dragDropStatus.HideDragImage(TRUE);
					SendMessage(WM_VSCROLL, SB_LINEDOWN, 0);
					dragDropStatus.ShowDragImage(TRUE);
				}
			}
		}
	}
}

BOOL ExplorerListView::IsInViewMode(ViewConstants view, BOOL detectExtendedTiles/* = TRUE*/)
{
	if(IsWindow()) {
		if(RunTimeHelper::IsCommCtrl6()) {
			if(!detectExtendedTiles) {
				if(view == vExtendedTiles) {
					view = vTiles;
				}
			}
			switch(view) {
				case vIcons:
					return (SendMessage(LVM_GETVIEW, 0, 0) == LV_VIEW_ICON);
					break;
				case vSmallIcons:
					return (SendMessage(LVM_GETVIEW, 0, 0) == LV_VIEW_SMALLICON);
					break;
				case vList:
					return (SendMessage(LVM_GETVIEW, 0, 0) == LV_VIEW_LIST);
					break;
				case vDetails:
					return (SendMessage(LVM_GETVIEW, 0, 0) == LV_VIEW_DETAILS);
					break;
				case vTiles:
					if(SendMessage(LVM_GETVIEW, 0, 0) == LV_VIEW_TILE) {
						if(detectExtendedTiles) {
							LVTILEVIEWINFO tileViewInfo = {0};
							tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
							tileViewInfo.dwMask = LVTVIM_TILESIZE;
							SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
							return ((tileViewInfo.dwFlags & LVTVIF_EXTENDED) == 0);
						} else {
							return TRUE;
						}
					} else {
						return FALSE;
					}
					break;
				case vExtendedTiles:
					if(SendMessage(LVM_GETVIEW, 0, 0) == LV_VIEW_TILE) {
						LVTILEVIEWINFO tileViewInfo = {0};
						tileViewInfo.cbSize = sizeof(LVTILEVIEWINFO);
						tileViewInfo.dwMask = LVTVIM_TILESIZE;
						SendMessage(LVM_GETTILEVIEWINFO, 0, reinterpret_cast<LPARAM>(&tileViewInfo));
						return ((tileViewInfo.dwFlags & LVTVIF_EXTENDED) == LVTVIF_EXTENDED);
						break;
					}
				default:
					return FALSE;
					break;
			}
		} else {
			switch(view) {
				case vIcons:
					return ((GetStyle() & (LVS_ICON | LVS_SMALLICON | LVS_LIST | LVS_REPORT)) == LVS_ICON);
					break;
				case vSmallIcons:
					return ((GetStyle() & (LVS_ICON | LVS_SMALLICON | LVS_LIST | LVS_REPORT)) == LVS_SMALLICON);
					break;
				case vList:
					return ((GetStyle() & (LVS_ICON | LVS_SMALLICON | LVS_LIST | LVS_REPORT)) == LVS_LIST);
					break;
				case vDetails:
					return ((GetStyle() & (LVS_ICON | LVS_SMALLICON | LVS_LIST | LVS_REPORT)) == LVS_REPORT);
					break;
				default:
					return FALSE;
					break;
			}
		}
	} else {
		return (properties.view == view);
	}
}

BOOL ExplorerListView::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL ExplorerListView::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}

void ExplorerListView::ReplaceGroupID(int oldGroupID, int newGroupID)
{
	#ifdef USE_STL
		std::vector<int>::iterator iter = std::find(groups.begin(), groups.end(), oldGroupID);
		if(iter != groups.end()) {
			*iter = newGroupID;
		}
	#else
		for(size_t i = 0; i < groups.GetCount(); ++i) {
			if(groups[i] == oldGroupID) {
				groups[i] = newGroupID;
				break;
			}
		}
	#endif
}

int ExplorerListView::GetUniqueGroupID(void)
{
	if(IsWindow()) {
		while(SendMessage(LVM_HASGROUP, nextGroupID, 0)) {
			// for some reason this ID is in use already
			++nextGroupID;
		}
		return nextGroupID;
	}
	return -1;
}


HRESULT ExplorerListView::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo = {0};
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = reinterpret_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}


BOOL ExplorerListView::IsComctl32Version600(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return ((major == 6) && (minor == 0));
	}
	return FALSE;
}

BOOL ExplorerListView::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}