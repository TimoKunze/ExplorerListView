// IPropertyValueImpl.cpp: An implementation of the IPropertyValue interface

#include "stdafx.h"
#include "IPropertyValueImpl.h"


HRESULT IPropertyValueImpl::CreateInstance(IUnknown* pOuter, REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(pOuter) {
		*ppInterfaceImpl = NULL;
		return CLASS_E_NOAGGREGATION;
	}

	// create and return the IPropertyValue object
	IPropertyValueImpl* pPropValue = new IPropertyValueImpl();
	pPropValue->AddRef();
	HRESULT hr = pPropValue->QueryInterface(requiredInterface, ppInterfaceImpl);
	pPropValue->Release();
	return hr;
}

HRESULT IPropertyValueImpl::CreateInstance(IPropertyValue** ppPropertyValue)
{
	return CreateInstance(NULL, IID_IPropertyValue, reinterpret_cast<LPVOID*>(ppPropertyValue));
}


//////////////////////////////////////////////////////////////////////
// implementation of IUnknown
HRESULT STDMETHODCALLTYPE IPropertyValueImpl::QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl)
{
	//const IID IID_IMultipleValues = {0xAD4AE633, 0x00BF, 0x4BB6, {0xB5, 0x27, 0x91, 0xA0, 0xE2, 0x8A, 0x6C, 0x78}};
	//const IID IID_IFilterCondition = {0xFCA2857D, 0x1760, 0x4AD3, {0x8C, 0x63, 0xC9, 0xB6, 0x02, 0xFC, 0xBA, 0xEA}};
	if(!ppInterfaceImpl) {
		return E_POINTER;
	}

	if(requiredInterface == IID_IUnknown) {
		*ppInterfaceImpl = this;
	} else if(requiredInterface == IID_IPropertyValue) {
		*ppInterfaceImpl = this;
	/*} else if(requiredInterface == IID_IMultipleValues) {
		// required by the sicMultiValueText control to support IPropertyValue::GetValue
		*ppInterfaceImpl = this;*/
	/*} else if(requiredInterface == IID_IFilterCondition) {
		*ppInterfaceImpl = this;*/
	} else {
		// the requested interface is not supported
		*ppInterfaceImpl = NULL;
		return E_NOINTERFACE;
	}

	// we return a new reference to this object, so increment the counter
	AddRef();
	return S_OK;
}

ULONG STDMETHODCALLTYPE IPropertyValueImpl::AddRef()
{
	return InterlockedIncrement(&properties.referenceCount);
}

ULONG STDMETHODCALLTYPE IPropertyValueImpl::Release()
{
	ULONG ret = InterlockedDecrement(&properties.referenceCount);
	if(!ret) {
		// the reference counter reached 0, so delete ourselves
		delete this;
	}
	return ret;
}
// implementation of IUnknown
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPropertyValue
HRESULT STDMETHODCALLTYPE IPropertyValueImpl::SetPropertyKey(PROPERTYKEY const& propKey)
{
	properties.propertyKey = propKey;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE IPropertyValueImpl::GetPropertyKey(PROPERTYKEY* pPropKey)
{
	if(!pPropKey) {
		return E_POINTER;
	}

	*pPropKey = properties.propertyKey;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE IPropertyValueImpl::GetValue(PROPVARIANT* pPropValue)
{
	if(!pPropValue) {
		return E_POINTER;
	}

	// NOTE: We'll crash on Windows 8 Release Preview if we call PropVariantClear for pPropValue.
	return PropVariantCopy(pPropValue, &properties.propertyValue);
}

HRESULT STDMETHODCALLTYPE IPropertyValueImpl::InitValue(PROPVARIANT const& propValue)
{
	PropVariantClear(&properties.propertyValue);
	return PropVariantCopy(&properties.propertyValue, &propValue);
}
// implementation of IPropertyValue
//////////////////////////////////////////////////////////////////////