// ListViewItems.cpp: Manages a collection of ListViewItem objects

#include "stdafx.h"
#include "ListViewItems.h"
#include "ClassFactory.h"


ListViewItems::ListViewItems()
{
	#ifdef UNICODE
		CLSIDFromString(OLESTR("{F8919B15-0236-4d2c-BDCA-3F0C832ACD8A}"), &clsidTILEVIEWSUBITEM);
	#else
		CLSIDFromString(OLESTR("{4D6B4D97-ED82-4234-9F68-225D8CDCEA9B}"), &clsidTILEVIEWSUBITEM);
	#endif
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ListViewItems::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IListViewItems, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ListViewItems::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ListViewItems>* pLvwItemsObj = NULL;
	CComObject<ListViewItems>::CreateInstance(&pLvwItemsObj);
	pLvwItemsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pLvwItemsObj->properties);

	pLvwItemsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLvwItemsObj->Release();
	return S_OK;
}

STDMETHODIMP ListViewItems::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	// check each item in the listview
	int numberOfItems = static_cast<int>(SendMessage(hWndLvw, LVM_GETITEMCOUNT, 0, 0));
	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		do {
			if(properties.lastEnumeratedItem.iItem >= 0) {
				if(properties.lastEnumeratedItem.iItem < numberOfItems) {
					properties.lastEnumeratedItem.iItem = GetNextItemToProcess(properties.lastEnumeratedItem.iItem, numberOfItems, hWndLvw);
				}
			} else {
				properties.lastEnumeratedItem.iItem = GetFirstItemToProcess(numberOfItems, hWndLvw);
			}
			if(properties.lastEnumeratedItem.iItem >= numberOfItems) {
				properties.lastEnumeratedItem.iItem = -1;
			}
		} while((properties.lastEnumeratedItem.iItem != -1) && (!IsPartOfCollection(properties.lastEnumeratedItem, hWndLvw)));

		if(properties.lastEnumeratedItem.iItem != -1) {
			ClassFactory::InitListItem(properties.lastEnumeratedItem, properties.pOwnerExLvw, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pItems[i].pdispVal));
			pItems[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP ListViewItems::Reset(void)
{
	properties.lastEnumeratedItem.iItem = -1;
	properties.lastEnumeratedItem.iGroup = 0;
	return S_OK;
}

STDMETHODIMP ListViewItems::Skip(ULONG numberOfItemsToSkip)
{
	VARIANT dummy;
	ULONG numItemsReturned = 0;
	// we could skip all items at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfItemsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numItemsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numItemsReturned == 0) {
			// there're no more items to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL ListViewItems::IsValidBooleanFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BOOL);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListViewItems::IsValidIntegerArrayFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					switch(element.vt) {
						case (VT_ARRAY | VT_UI1):
						case (VT_ARRAY | VT_I1):
						case (VT_ARRAY | VT_UI2):
						case (VT_ARRAY | VT_I2):
						case (VT_ARRAY | VT_UI4):
						case (VT_ARRAY | VT_I4):
						case (VT_ARRAY | VT_UINT):
						case (VT_ARRAY | VT_INT):
							// the element is valid
							break;
						case (VT_ARRAY | VT_VARIANT): {
							LONG innerLBound = 0;
							LONG innerUBound = 0;
							if((SafeArrayGetLBound(element.parray, 1, &innerLBound) == S_OK) && (SafeArrayGetUBound(element.parray, 1, &innerUBound) == S_OK)) {
								VARIANT buffer;
								for(LONG j = innerLBound; j <= innerUBound && isValid; ++j) {
									if(SafeArrayGetElement(element.parray, &j, &buffer) == S_OK) {
										isValid = SUCCEEDED(VariantChangeType(&buffer, &buffer, 0, VT_INT));
										VariantClear(&buffer);
									} else {
										isValid = FALSE;
									}
								}
							} else {
								isValid = FALSE;
							}
							break;
						}
						default:
							// the element is invalid - abort
							isValid = FALSE;
							break;
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListViewItems::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListViewItems::IsValidIntegerFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4));
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListViewItems::IsValidListViewGroupFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					if(element.vt == VT_DISPATCH && element.pdispVal) {
						CComQIPtr<IListViewGroup, &IID_IListViewGroup> pLvwGroup(element.pdispVal);
						if(!pLvwGroup) {
							// the element is invalid - abort
							isValid = FALSE;
						}
					} else {
						// we also support group ids
						switch(element.vt) {
							case VT_UI1:
							case VT_I1:
							case VT_UI2:
							case VT_I2:
							case VT_UI4:
							case VT_I4:
							case VT_UINT:
							case VT_INT:
								// the element is valid
								break;
							default:
								// the element is invalid - abort
								isValid = FALSE;
								break;
						}
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListViewItems::IsValidListViewWorkAreaFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					if(element.vt == VT_DISPATCH && element.pdispVal) {
						CComQIPtr<IListViewWorkArea, &IID_IListViewWorkArea> pLvwWorkArea(element.pdispVal);
						if(!pLvwWorkArea) {
							// the element is invalid - abort
							isValid = FALSE;
						}
					} else {
						// we also support working area indexes
						switch(element.vt) {
							case VT_UI1:
							case VT_I1:
							case VT_UI2:
							case VT_I2:
							case VT_UI4:
							case VT_I4:
							case VT_UINT:
							case VT_INT:
								// the element is valid
								break;
							default:
								// the element is invalid - abort
								isValid = FALSE;
								break;
						}
					}

					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ListViewItems::IsValidStringFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BSTR);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

int ListViewItems::GetFirstItemToProcess(int numberOfItems, HWND hWndLvw)
{
	ATLASSERT(IsWindow(hWndLvw));

	if(numberOfItems == 0) {
		return -1;
	}

	UINT flags = LVNI_ALL;
	if((properties.filterType[fpSelected] == ftIncluding) && !properties.comparisonFunction[fpSelected]) {
		flags |= LVNI_SELECTED;
	}
	if((properties.filterType[fpGhosted] == ftIncluding) && !properties.comparisonFunction[fpGhosted]) {
		flags |= LVNI_CUT;
	}
	if(flags == LVNI_ALL) {
		return 0;
	} else {
		return static_cast<int>(SendMessage(hWndLvw, LVM_GETNEXTITEM, static_cast<WPARAM>(-1), MAKELPARAM(flags, 0)));
	}
}

int ListViewItems::GetNextItemToProcess(int previousItem, int numberOfItems, HWND hWndLvw)
{
	ATLASSERT(IsWindow(hWndLvw));

	UINT flags = LVNI_ALL;
	if((properties.filterType[fpSelected] == ftIncluding) && !properties.comparisonFunction[fpSelected]) {
		flags |= LVNI_SELECTED;
	}
	if((properties.filterType[fpGhosted] == ftIncluding) && !properties.comparisonFunction[fpGhosted]) {
		flags |= LVNI_CUT;
	}
	if(flags == 0) {
		if(previousItem < numberOfItems - 1) {
			return previousItem + 1;
		} else {
			return -1;
		}
	} else {
		return static_cast<int>(SendMessage(hWndLvw, LVM_GETNEXTITEM, previousItem, MAKELPARAM(flags, 0)));
	}
}

BOOL ListViewItems::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(VARIANT_BOOL, VARIANT_BOOL);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.boolVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(element.boolVal == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListViewItems::IsIntegerArrayInSafeArray(LPSAFEARRAY pSafeArray, PINT pArray, int elements, LPVOID pComparisonFunction)
{
	LPSAFEARRAY pValue = NULL;
	if(pComparisonFunction) {
		pValue = SafeArrayCreateVectorEx(VT_I4, 1, elements, NULL);
		for(LONG j = 1; j <= elements; ++j) {
			SafeArrayPutElement(pValue, &j, &pArray[j - 1]);
		}
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LPSAFEARRAY, LPSAFEARRAY);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(pValue, element.parray)) {
				foundMatch = TRUE;
			}
		} else {
			LONG innerLBound = 0;
			LONG innerUBound = 0;
			SafeArrayGetLBound(element.parray, 1, &innerLBound);
			SafeArrayGetUBound(element.parray, 1, &innerUBound);

			if((innerUBound - innerLBound + 1) == elements) {
				foundMatch = TRUE;
				VARTYPE vt = VT_EMPTY;
				SafeArrayGetVartype(element.parray, &vt);
				for(LONG j = innerLBound; j <= innerUBound && foundMatch; ++j) {
					INT value = 0;
					if(vt == VT_VARIANT) {
						VARIANT v;
						SafeArrayGetElement(element.parray, &j, &v);
						if(SUCCEEDED(VariantChangeType(&v, &v, 0, VT_INT))) {
							value = v.intVal;
						}
						VariantClear(&v);
					} else if(vt == VT_I1 || vt == VT_UI1 || vt == VT_I2 || vt == VT_UI2 || vt == VT_I4 || vt == VT_UI4 || vt == VT_INT || vt == VT_UINT) {
						SafeArrayGetElement(element.parray, &j, &value);
					}
					if(value != pArray[j - innerLBound]) {
						foundMatch = FALSE;
					}
				}
			}
		}
		VariantClear(&element);
	}

	if(pValue) {
		SafeArrayDestroy(pValue);
	}

	return foundMatch;
}

BOOL ListViewItems::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			v = element.intVal;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListViewItems::IsListViewGroupInSafeArray(LPSAFEARRAY pSafeArray, int groupID, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int elementGroupID = -1;
		if(element.vt == VT_DISPATCH) {
			CComQIPtr<IListViewGroup, &IID_IListViewGroup> pLvwGroup(element.pdispVal);
			if(pLvwGroup) {
				LONG l = -1;
				pLvwGroup->get_ID(&l);
				elementGroupID = l;
			}
		} else if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			elementGroupID = element.intVal;
		}

		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(IListViewGroup*, IListViewGroup*);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			CComPtr<IListViewGroup> pLvwGroup1 = ClassFactory::InitGroup(groupID, properties.pOwnerExLvw);
			CComPtr<IListViewGroup> pLvwGroup2 = ClassFactory::InitGroup(elementGroupID, properties.pOwnerExLvw);
			if(pComparisonFn(pLvwGroup1, pLvwGroup2)) {
				foundMatch = TRUE;
			}
		} else {
			if(elementGroupID == groupID) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListViewItems::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.bstrVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(properties.caseSensitiveFilters) {
				if(lstrcmpW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			} else {
				if(lstrcmpiW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ListViewItems::IsPartOfCollection(LVITEMINDEX& itemIndex, HWND hWndLvw/* = NULL*/)
{
	if(!hWndLvw) {
		hWndLvw = properties.GetExLvwHWnd();
	}
	ATLASSERT(IsWindow(hWndLvw));

	if(!IsValidItemIndex(itemIndex.iItem, hWndLvw)) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	LVITEM item = {0};

	if(properties.filterType[fpIndex] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpIndex].parray, itemIndex.iItem, properties.comparisonFunction[fpIndex])) {
			if(properties.filterType[fpIndex] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpIndex] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[fpWorkArea] != ftDeactivated) {
		int workAreaIndex = -1;
		CComPtr<IListViewItem> pLvwItem = ClassFactory::InitListItem(itemIndex, properties.pOwnerExLvw);
		if(pLvwItem) {
			CComPtr<IListViewWorkArea> pLvwWorkArea = NULL;
			pLvwItem->get_WorkArea(&pLvwWorkArea);
			if(pLvwWorkArea) {
				LONG l = -1;
				pLvwWorkArea->get_Index(&l);
				workAreaIndex = l;
			}
			pLvwWorkArea = NULL;
		}
		pLvwItem = NULL;
		if(workAreaIndex == -1) {
			goto Exit;
		}

		LONG lBound = 0;
		LONG uBound = 0;
		SafeArrayGetLBound(properties.filter[fpWorkArea].parray, 1, &lBound);
		SafeArrayGetUBound(properties.filter[fpWorkArea].parray, 1, &uBound);
		VARIANT element;
		BOOL foundMatch = FALSE;
		for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
			SafeArrayGetElement(properties.filter[fpWorkArea].parray, &i, &element);
			int indexToCompareWith = -1;
			if(element.vt == VT_DISPATCH) {
				CComQIPtr<IListViewWorkArea, &IID_IListViewWorkArea> pLvwWorkArea(element.pdispVal);
				if(pLvwWorkArea) {
					LONG l = -1;
					pLvwWorkArea->get_Index(&l);
					indexToCompareWith = l;
				}
			} else if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
				indexToCompareWith = element.intVal;
			}
			if(properties.comparisonFunction[fpWorkArea]) {
				typedef BOOL WINAPI ComparisonFn(IListViewWorkArea*, IListViewWorkArea*);
				ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(properties.comparisonFunction[fpWorkArea]);
				CComPtr<IListViewWorkArea> pLvwWorkArea1 = ClassFactory::InitWorkArea(workAreaIndex, properties.pOwnerExLvw);
				CComPtr<IListViewWorkArea> pLvwWorkArea2 = ClassFactory::InitWorkArea(indexToCompareWith, properties.pOwnerExLvw);
				if(pComparisonFn(pLvwWorkArea1, pLvwWorkArea2)) {
					foundMatch = TRUE;
				}
			} else {
				if(workAreaIndex == indexToCompareWith) {
					foundMatch = TRUE;
				}
			}
			VariantClear(&element);
		}

		if(foundMatch) {
			if(properties.filterType[fpWorkArea] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpWorkArea] == ftIncluding) {
				goto Exit;
			}
		}
	}

	item.iItem = itemIndex.iItem;
	item.iGroup = itemIndex.iGroup;
	BOOL mustRetrieveItemData = FALSE;
	if(properties.filterType[fpActivating] != ftDeactivated) {
		item.mask |= LVIF_STATE;
		item.stateMask |= LVIS_ACTIVATING;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpGhosted] != ftDeactivated) {
		item.mask |= LVIF_STATE;
		item.stateMask |= LVIS_CUT;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpGlowing] != ftDeactivated) {
		item.mask |= LVIF_STATE;
		item.stateMask |= LVIS_GLOW;
		mustRetrieveItemData = TRUE;
	}
	if((properties.filterType[fpGroup] != ftDeactivated) && RunTimeHelper::IsCommCtrl6()) {
		item.mask |= LVIF_GROUPID;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpIconIndex] != ftDeactivated) {
		item.mask |= LVIF_IMAGE;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpIndent] != ftDeactivated) {
		item.mask |= LVIF_INDENT;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpItemData] != ftDeactivated) {
		item.mask |= LVIF_PARAM;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpOverlayIndex] != ftDeactivated) {
		item.mask |= LVIF_STATE;
		item.stateMask |= LVIS_OVERLAYMASK;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpSelected] != ftDeactivated) {
		item.mask |= LVIF_STATE;
		item.stateMask |= LVIS_SELECTED;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpStateImageIndex] != ftDeactivated) {
		item.mask |= LVIF_STATE;
		item.stateMask |= LVIS_STATEIMAGEMASK;
		mustRetrieveItemData = TRUE;
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		item.mask |= LVIF_TEXT;
		item.cchTextMax = MAX_ITEMTEXTLENGTH;
		item.pszText = reinterpret_cast<LPTSTR>(malloc((item.cchTextMax + 1) * sizeof(TCHAR)));
		mustRetrieveItemData = TRUE;
	}
	if((properties.filterType[fpTileViewColumns] != ftDeactivated) && RunTimeHelper::IsCommCtrl6()) {
		item.mask |= LVIF_COLUMNS;
		item.puColumns = new UINT[200];
		if(properties.pOwnerExLvw->IsComctl32Version610OrNewer()) {
			item.cColumns = 200;
		}
		mustRetrieveItemData = TRUE;
	}

	if(mustRetrieveItemData) {
		if(!SendMessage(hWndLvw, LVM_GETITEM, 0, reinterpret_cast<LPARAM>(&item))) {
			goto Exit;
		}

		// apply the filters

		if(properties.filterType[fpSelected] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpSelected].parray, BOOL2VARIANTBOOL((item.state & LVIS_SELECTED) == LVIS_SELECTED), properties.comparisonFunction[fpSelected])) {
				if(properties.filterType[fpSelected] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSelected] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpStateImageIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpStateImageIndex].parray, (item.state & LVIS_STATEIMAGEMASK) >> 12, properties.comparisonFunction[fpStateImageIndex])) {
				if(properties.filterType[fpStateImageIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpStateImageIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpItemData] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpItemData].parray, static_cast<int>(item.lParam), properties.comparisonFunction[fpItemData])) {
				if(properties.filterType[fpItemData] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpItemData] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if((properties.filterType[fpGroup] != ftDeactivated) && RunTimeHelper::IsCommCtrl6()) {
			if(IsListViewGroupInSafeArray(properties.filter[fpGroup].parray, item.iGroupId, properties.comparisonFunction[fpGroup])) {
				if(properties.filterType[fpGroup] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpGroup] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpActivating] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpActivating].parray, BOOL2VARIANTBOOL((item.state & LVIS_ACTIVATING) == LVIS_ACTIVATING), properties.comparisonFunction[fpActivating])) {
				if(properties.filterType[fpActivating] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpActivating] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpGhosted] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpGhosted].parray, BOOL2VARIANTBOOL((item.state & LVIS_CUT) == LVIS_CUT), properties.comparisonFunction[fpGhosted])) {
				if(properties.filterType[fpGhosted] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpGhosted] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpGlowing] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpGlowing].parray, BOOL2VARIANTBOOL((item.state & LVIS_GLOW) == LVIS_GLOW), properties.comparisonFunction[fpGlowing])) {
				if(properties.filterType[fpGlowing] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpGlowing] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpOverlayIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpOverlayIndex].parray, (item.state & LVIS_OVERLAYMASK) >> 8, properties.comparisonFunction[fpOverlayIndex])) {
				if(properties.filterType[fpOverlayIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpOverlayIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIconIndex] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIconIndex].parray, item.iImage, properties.comparisonFunction[fpIconIndex])) {
				if(properties.filterType[fpIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIndent] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIndent].parray, item.iIndent, properties.comparisonFunction[fpIndent])) {
				if(properties.filterType[fpIndent] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIndent] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if((properties.filterType[fpTileViewColumns] != ftDeactivated) && RunTimeHelper::IsCommCtrl6()) {
			if(IsIntegerArrayInSafeArray(properties.filter[fpTileViewColumns].parray, reinterpret_cast<PINT>(item.puColumns), item.cColumns, properties.comparisonFunction[fpTileViewColumns])) {
				if(properties.filterType[fpTileViewColumns] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpTileViewColumns] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpText] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpText].parray, CComBSTR(item.pszText), properties.comparisonFunction[fpText])) {
				if(properties.filterType[fpText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpText] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	if(item.pszText) {
		SECUREFREE(item.pszText);
	}
	if(item.puColumns) {
		delete[] item.puColumns;
	}
	if(item.piColFmt) {
		delete[] item.piColFmt;
	}
	return isPart;
}

#ifdef INCLUDESHELLBROWSERINTERFACE
	int ListViewItems::Add(LPTSTR pItemText, int insertAt, int iconIndex, int overlayIndex, int itemIndentation, LPARAM itemData, BOOL isGhosted, int groupID, VARIANT* pTileViewColumns, BOOL setShellItemFlag)
#else
	int ListViewItems::Add(LPTSTR pItemText, int insertAt, int iconIndex, int overlayIndex, int itemIndentation, LPARAM itemData, BOOL isGhosted, int groupID, VARIANT* pTileViewColumns)
#endif
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	int itemIndex = -1;
	LVITEM insertionData = {0};
	insertionData.iItem = insertAt;
	insertionData.lParam = itemData;
	insertionData.iIndent = itemIndentation;
	insertionData.mask = LVIF_INDENT | LVIF_PARAM | LVIF_TEXT;

	if(iconIndex != -2) {
		insertionData.iImage = iconIndex;
		insertionData.mask |= LVIF_IMAGE;
	}
	if(overlayIndex > 0) {
		insertionData.state |= INDEXTOOVERLAYMASK(overlayIndex);
		insertionData.stateMask |= LVIS_OVERLAYMASK;
		insertionData.mask |= LVIF_STATE;
	}
	if(isGhosted) {
		insertionData.state |= LVIS_CUT;
		insertionData.stateMask |= LVIS_CUT;
		insertionData.mask |= LVIF_STATE;
	}
	if(RunTimeHelper::IsCommCtrl6()) {
		insertionData.iGroupId = groupID;
		insertionData.mask |= LVIF_GROUPID;

		if(pTileViewColumns->vt == VT_ERROR && pTileViewColumns->scode == DISP_E_PARAMNOTFOUND) {
			// the argument wasn't specified
			insertionData.cColumns = 0;
		} else if(pTileViewColumns->vt == VT_EMPTY) {
			insertionData.cColumns = I_COLUMNSCALLBACK;
			insertionData.mask |= LVIF_COLUMNS | LVIF_COLFMT;
		} else if(pTileViewColumns->vt & VT_ARRAY) {
			// an array
			if(pTileViewColumns->parray) {
				LONG l = 0;
				SafeArrayGetLBound(pTileViewColumns->parray, 1, &l);
				LONG u = 0;
				SafeArrayGetUBound(pTileViewColumns->parray, 1, &u);
				VARTYPE vt = 0;
				SafeArrayGetVartype(pTileViewColumns->parray, &vt);

				insertionData.cColumns = u - l + 1;
				insertionData.puColumns = new UINT[insertionData.cColumns];
				insertionData.piColFmt = new int[insertionData.cColumns];

				TILEVIEWSUBITEM element = {0};
				for(LONG i = l; i <= u; ++i) {
					if(vt == VT_RECORD) {
						ATLVERIFY(SUCCEEDED(SafeArrayGetElement(pTileViewColumns->parray, &i, &element)));
					} else {
						VARIANT v;
						ATLVERIFY(SUCCEEDED(SafeArrayGetElement(pTileViewColumns->parray, &i, &v)));
						element = *reinterpret_cast<TILEVIEWSUBITEM*>(v.pvRecord);
						VariantClear(&v);
					}
					insertionData.puColumns[i - l] = element.ColumnIndex;
					insertionData.piColFmt[i - l] = 0;
					if(element.BeginNewColumn != VARIANT_FALSE) {
						insertionData.piColFmt[i - l] |= LVCFMT_LINE_BREAK;
					}
					if(element.FillRemainder != VARIANT_FALSE) {
						insertionData.piColFmt[i - l] |= LVCFMT_FILL;
					}
					if(element.WrapToMultipleLines != VARIANT_FALSE) {
						insertionData.piColFmt[i - l] |= LVCFMT_WRAP;
					}
					if(element.NoTitle != VARIANT_FALSE) {
						insertionData.piColFmt[i - l] |= LVCFMT_NO_TITLE;
					}
				}
			} else {
				// an empty array
				insertionData.cColumns = 0;
			}
			insertionData.mask |= LVIF_COLUMNS | LVIF_COLFMT;
		} else if(pTileViewColumns->vt == VT_RECORD) {
			// a single value
			TILEVIEWSUBITEM* pData = reinterpret_cast<TILEVIEWSUBITEM*>(pTileViewColumns->pvRecord);
			ATLASSERT_POINTER(pData, TILEVIEWSUBITEM);

			insertionData.cColumns = 1;
			insertionData.puColumns = new UINT[insertionData.cColumns];
			insertionData.piColFmt = new int[insertionData.cColumns];

			insertionData.puColumns[0] = pData->ColumnIndex;
			insertionData.piColFmt[0] = 0;
			if(pData->BeginNewColumn != VARIANT_FALSE) {
				insertionData.piColFmt[0] |= LVCFMT_LINE_BREAK;
			}
			if(pData->FillRemainder != VARIANT_FALSE) {
				insertionData.piColFmt[0] |= LVCFMT_FILL;
			}
			if(pData->WrapToMultipleLines != VARIANT_FALSE) {
				insertionData.piColFmt[0] |= LVCFMT_WRAP;
			}
			if(pData->NoTitle != VARIANT_FALSE) {
				insertionData.piColFmt[0] |= LVCFMT_NO_TITLE;
			}
			insertionData.mask |= LVIF_COLUMNS | LVIF_COLFMT;
		}
	}

	#ifdef INCLUDESHELLBROWSERINTERFACE
		if(setShellItemFlag) {
			insertionData.mask |= LVIF_ISASHELLITEM;
		}
	#endif

	insertionData.pszText = pItemText;
	itemIndex = static_cast<int>(SendMessage(hWndLvw, LVM_INSERTITEM, 0, reinterpret_cast<LPARAM>(&insertionData)));
	if(itemIndex >= 0 && (insertionData.mask & (LVIF_COLUMNS | LVIF_COLFMT))) {
		// small hack to make LVIF_COLFMT really work
		insertionData.mask = LVIF_COLUMNS | LVIF_COLFMT;
		insertionData.iItem = itemIndex;
		insertionData.iGroup = 0;
		SendMessage(hWndLvw, LVM_SETITEM, 0, reinterpret_cast<LPARAM>(&insertionData));
	}

	if(insertionData.puColumns) {
		delete[] insertionData.puColumns;
	}
	if(insertionData.piColFmt) {
		delete[] insertionData.piColFmt;
	}
	return itemIndex;
}

void ListViewItems::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(!(filteredProperty == fpActivating || filteredProperty == fpGhosted || filteredProperty == fpGlowing || filteredProperty == fpSelected)) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; i <= uBound; ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if(numberOfTrues > 0 && numberOfFalses > 0) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if(numberOfTrues == 0 && numberOfFalses > 1) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if(numberOfFalses == 0 && numberOfTrues > 1) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}

#ifdef USE_STL
	HRESULT ListViewItems::RemoveItems(std::list<int>& itemsToRemove, HWND hWndLvw)
#else
	HRESULT ListViewItems::RemoveItems(CAtlList<int>& itemsToRemove, HWND hWndLvw)
#endif
{
	ATLASSERT(IsWindow(hWndLvw));

	CWindowEx2(hWndLvw).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		itemsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = itemsToRemove.begin(); iter != itemsToRemove.end(); ++iter) {
			SendMessage(hWndLvw, LVM_DELETEITEM, *iter, 0);
		}
	#else
		// perform a crude bubble sort
		for(size_t j = 0; j < itemsToRemove.GetCount(); ++j) {
			for(size_t i = 0; i < itemsToRemove.GetCount() - 1; ++i) {
				if(itemsToRemove.GetAt(itemsToRemove.FindIndex(i)) < itemsToRemove.GetAt(itemsToRemove.FindIndex(i + 1))) {
					itemsToRemove.SwapElements(itemsToRemove.FindIndex(i), itemsToRemove.FindIndex(i + 1));
				}
			}
		}

		for(size_t i = 0; i < itemsToRemove.GetCount(); ++i) {
			SendMessage(hWndLvw, LVM_DELETEITEM, itemsToRemove.GetAt(itemsToRemove.FindIndex(i)), 0);
		}
	#endif
	CWindowEx2(hWndLvw).InternalSetRedraw(TRUE);

	return S_OK;
}


ListViewItems::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerExLvw) {
		pOwnerExLvw->Release();
	}
}

void ListViewItems::Properties::CopyTo(ListViewItems::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerExLvw = this->pOwnerExLvw;
		if(pOwnerExLvw) {
			pOwnerExLvw->AddRef();
		}
		pTarget->lastEnumeratedItem = this->lastEnumeratedItem;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND ListViewItems::Properties::GetExLvwHWnd(void)
{
	ATLASSUME(pOwnerExLvw);

	OLE_HANDLE handle = NULL;
	pOwnerExLvw->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ListViewItems::SetOwner(ExplorerListView* pOwner)
{
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->Release();
	}
	properties.pOwnerExLvw = pOwner;
	if(properties.pOwnerExLvw) {
		properties.pOwnerExLvw->AddRef();
	}
}


STDMETHODIMP ListViewItems::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ListViewItems::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ListViewItems::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListViewItems::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListViewItems::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		VariantClear(pValue);
		VariantCopy(pValue, &properties.filter[filteredProperty]);
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListViewItems::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		// check 'newValue'
		switch(filteredProperty) {
			case fpActivating:
			case fpGhosted:
			case fpGlowing:
			case fpSelected:
				if(!IsValidBooleanFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpGroup:
				if(!IsValidListViewGroupFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpIconIndex:
				if(!IsValidIntegerFilter(newValue, -1)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpIndent:
			case fpIndex:
			case fpOverlayIndex:
			case fpStateImageIndex:
				if(!IsValidIntegerFilter(newValue, 0)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpItemData:
				if(!IsValidIntegerFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpText:
				if(!IsValidStringFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpTileViewColumns:
				if(!IsValidIntegerArrayFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
			case fpWorkArea:
				if(!IsValidListViewWorkAreaFilter(newValue)) {
					// invalid value - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				}
				break;
		}

		VariantClear(&properties.filter[filteredProperty]);
		VariantCopy(&properties.filter[filteredProperty], &newValue);
		OptimizeFilter(filteredProperty);
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListViewItems::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		*pValue = properties.filterType[filteredProperty];
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListViewItems::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(filteredProperty >= 0 && filteredProperty < NUMBEROFFILTERS) {
		properties.filterType[filteredProperty] = newValue;
		if(newValue != ftDeactivated) {
			properties.usingFilters = TRUE;
		} else {
			properties.usingFilters = FALSE;
			for(int i = 0; i < NUMBEROFFILTERS; ++i) {
				if(properties.filterType[i] != ftDeactivated) {
					properties.usingFilters = TRUE;
					break;
				}
			}
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ListViewItems::get_Item(LONG itemIdentifier, LONG groupIndex/* = 0*/, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, IListViewItem** ppItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppItem, IListViewItem*);
	if(!ppItem) {
		return E_POINTER;
	}

	// retrieve the item's index
	int itemIndex = -2;
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerExLvw) {
				itemIndex = properties.pOwnerExLvw->IDToItemIndex(itemIdentifier);
				if(itemIndex == -1) {
					itemIndex = -2;
				}
			}
			break;
		case iitIndex:
			itemIndex = itemIdentifier;
			break;
	}
	if(itemIndex == -1) {
		if(properties.usingFilters) {
			// we don't support an index of -1 if we're using filters
			itemIndex = -2;
		}
	}

	if(itemIndex > -2) {
		LVITEMINDEX i = {itemIndex, groupIndex};
		if(itemIndex == -1) {
			ClassFactory::InitListItem(i, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem), FALSE);
		} else if(IsPartOfCollection(i)) {
			ClassFactory::InitListItem(i, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppItem));
		}
		return S_OK;
	} else {
		// item not found
		if(itemIdentifierType == iitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}
}

STDMETHODIMP ListViewItems::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ListViewItems::Add(BSTR itemText, LONG insertAt/* = -1*/, LONG iconIndex/* = -2*/, LONG itemIndentation/* = 0*/, LONG itemData/* = 0*/, LONG groupID/* = -2*/, VARIANT tileViewColumns/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IListViewItem** ppAddedItem/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedItem, IListViewItem*);
	if(!ppAddedItem) {
		return E_POINTER;
	}
	if(insertAt == -1) {
		insertAt = 0x7FFFFFFF;
	}
	ATLASSERT(insertAt >= 0);
	if(insertAt < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HRESULT hr = E_FAIL;

	if(tileViewColumns.vt == VT_ERROR && tileViewColumns.scode == DISP_E_PARAMNOTFOUND) {
		// all fine, do nothing
	} else if(tileViewColumns.vt == VT_EMPTY) {
		// all fine, do nothing
	} else if(tileViewColumns.vt & VT_ARRAY) {
		// an array
		if(tileViewColumns.parray) {
			LONG l = 0;
			SafeArrayGetLBound(tileViewColumns.parray, 1, &l);
			LONG u = 0;
			SafeArrayGetUBound(tileViewColumns.parray, 1, &u);
			if(u < l) {
				// invalid arg - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}

			VARTYPE vt = 0;
			SafeArrayGetVartype(tileViewColumns.parray, &vt);
			if(vt == VT_VARIANT) {
				for(LONG i = l; i <= u; ++i) {
					VARIANT buffer;
					SafeArrayGetElement(tileViewColumns.parray, &i, &buffer);
					if(buffer.vt != VT_RECORD) {
						// invalid arg - raise VB runtime error 380
						VariantClear(&buffer);
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					if(!buffer.pRecInfo) {
						// invalid arg - raise VB runtime error 380
						VariantClear(&buffer);
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					GUID recordGuid = {0};
					ATLVERIFY(SUCCEEDED(buffer.pRecInfo->GetGuid(&recordGuid)));
					ULONG elementSize = 0;
					ATLVERIFY(SUCCEEDED(buffer.pRecInfo->GetSize(&elementSize)));
					if(recordGuid != clsidTILEVIEWSUBITEM || elementSize > sizeof(TILEVIEWSUBITEM)) {
						// invalid arg - raise VB runtime error 380
						VariantClear(&buffer);
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
					VariantClear(&buffer);
				}
			} else if(vt == VT_RECORD) {
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				ATLVERIFY(SUCCEEDED(SafeArrayGetRecordInfo(tileViewColumns.parray, &pRecordInfo)));
				if(!pRecordInfo) {
					// invalid arg - raise VB runtime error 380
					return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
				} else {
					GUID recordGuid = {0};
					ATLVERIFY(SUCCEEDED(pRecordInfo->GetGuid(&recordGuid)));
					ULONG elementSize = 0;
					ATLVERIFY(SUCCEEDED(pRecordInfo->GetSize(&elementSize)));
					if(recordGuid != clsidTILEVIEWSUBITEM || elementSize > sizeof(TILEVIEWSUBITEM)) {
						// invalid arg - raise VB runtime error 380
						return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
					}
				}
			} else {
				// invalid arg - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
		} else {
			// an empty array - all fine, do nothing
		}
	} else {
		// a single value
		if(tileViewColumns.vt != VT_RECORD) {
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		if(!tileViewColumns.pRecInfo) {
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		GUID recordGuid = {0};
		ATLVERIFY(SUCCEEDED(tileViewColumns.pRecInfo->GetGuid(&recordGuid)));
		ULONG elementSize = 0;
		ATLVERIFY(SUCCEEDED(tileViewColumns.pRecInfo->GetSize(&elementSize)));
		if(recordGuid != clsidTILEVIEWSUBITEM || elementSize > sizeof(TILEVIEWSUBITEM)) {
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	}

	LPTSTR pItemText;
	#ifndef UNICODE
		COLE2T converter(itemText);
	#endif
	if(itemText) {
		#ifdef UNICODE
			pItemText = OLE2W(itemText);
		#else
			pItemText = converter;
		#endif
	} else {
		pItemText = LPSTR_TEXTCALLBACK;
	}

	// finally insert the item
	*ppAddedItem = NULL;
	#ifdef INCLUDESHELLBROWSERINTERFACE
		LVITEMINDEX itemIndex = {Add(pItemText, insertAt, iconIndex, 0, itemIndentation, itemData, FALSE, groupID, &tileViewColumns, FALSE), 0};
	#else
		LVITEMINDEX itemIndex = {Add(pItemText, insertAt, iconIndex, 0, itemIndentation, itemData, FALSE, groupID, &tileViewColumns), 0};
	#endif
	if(itemIndex.iItem != -1) {
		ClassFactory::InitListItem(itemIndex, properties.pOwnerExLvw, IID_IListViewItem, reinterpret_cast<LPUNKNOWN*>(ppAddedItem), FALSE);
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ListViewItems::Contains(LONG itemIdentifier, LONG groupIndex/* = 0*/, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the item's index
	LVITEMINDEX itemIndex = {-1, groupIndex};
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerExLvw) {
				itemIndex.iItem = properties.pOwnerExLvw->IDToItemIndex(itemIdentifier);
			}
			break;
		case iitIndex:
			itemIndex.iItem = itemIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(itemIndex));
	return S_OK;
}

STDMETHODIMP ListViewItems::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(!properties.usingFilters) {
		*pValue = static_cast<LONG>(SendMessage(hWndLvw, LVM_GETITEMCOUNT, 0, 0));
		return S_OK;
	}

	if(properties.filterType[fpSelected] != ftDeactivated && !properties.comparisonFunction[fpSelected]) {
		int numberOfActiveFilters = 0;
		for(int i = 0; i < NUMBEROFFILTERS; ++i) {
			if(properties.filterType[i] != ftDeactivated) {
				++numberOfActiveFilters;
			}
		}
		if(numberOfActiveFilters == 1) {
			// we can use LVM_GETSELECTEDCOUNT
			if(properties.filterType[fpSelected] == ftIncluding) {
				*pValue = static_cast<LONG>(SendMessage(hWndLvw, LVM_GETSELECTEDCOUNT, 0, 0));
				return S_OK;
			} else {
				*pValue = static_cast<LONG>(SendMessage(hWndLvw, LVM_GETITEMCOUNT, 0, 0)) - static_cast<LONG>(SendMessage(hWndLvw, LVM_GETSELECTEDCOUNT, 0, 0));
				return S_OK;
			}
		}
	}

	// count the items "by hand"
	*pValue = 0;
	int numberOfItems = static_cast<int>(SendMessage(hWndLvw, LVM_GETITEMCOUNT, 0, 0));
	LVITEMINDEX itemIndex = {GetFirstItemToProcess(numberOfItems, hWndLvw), 0};
	while(itemIndex.iItem != -1) {
		if(IsPartOfCollection(itemIndex, hWndLvw)) {
			++(*pValue);
		}
		itemIndex.iItem = GetNextItemToProcess(itemIndex.iItem, numberOfItems, hWndLvw);
	}
	return S_OK;
}

STDMETHODIMP ListViewItems::Remove(LONG itemIdentifier, ItemIdentifierTypeConstants itemIdentifierType/* = iitIndex*/)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	// retrieve the item's index
	LVITEMINDEX itemIndex = {-1, 0};
	switch(itemIdentifierType) {
		case iitID:
			if(properties.pOwnerExLvw) {
				itemIndex.iItem = properties.pOwnerExLvw->IDToItemIndex(itemIdentifier);
			}
			break;
		case iitIndex:
			itemIndex.iItem = itemIdentifier;
			break;
	}

	if(itemIndex.iItem != -1) {
		if(IsPartOfCollection(itemIndex)) {
			if(SendMessage(hWndLvw, LVM_DELETEITEM, itemIndex.iItem, 0)) {
				return S_OK;
			}
		}
	} else {
		// item not found
		if(itemIdentifierType == iitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ListViewItems::RemoveAll(void)
{
	HWND hWndLvw = properties.GetExLvwHWnd();
	ATLASSERT(IsWindow(hWndLvw));

	if(!properties.usingFilters) {
		if(SendMessage(hWndLvw, LVM_DELETEALLITEMS, 0, 0)) {
			return S_OK;
		} else {
			return E_FAIL;
		}
	}

	// find the items to remove manually
	#ifdef USE_STL
		std::list<int> itemsToRemove;
	#else
		CAtlList<int> itemsToRemove;
	#endif
	int numberOfItems = static_cast<int>(SendMessage(hWndLvw, LVM_GETITEMCOUNT, 0, 0));
	LVITEMINDEX itemIndex = {GetFirstItemToProcess(numberOfItems, hWndLvw), 0};
	while(itemIndex.iItem != -1) {
		if(IsPartOfCollection(itemIndex, hWndLvw)) {
			#ifdef USE_STL
				itemsToRemove.push_back(itemIndex.iItem);
			#else
				itemsToRemove.AddTail(itemIndex.iItem);
			#endif
		}
		itemIndex.iItem = GetNextItemToProcess(itemIndex.iItem, numberOfItems, hWndLvw);
	}
	return RemoveItems(itemsToRemove, hWndLvw);
}