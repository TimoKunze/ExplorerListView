//////////////////////////////////////////////////////////////////////
/// \class IPropertyValue
/// \author Timo "TimoSoft" Kunze, Microsoft
/// \brief <em>Interface for accessing a property</em>
///
/// This interface is used to manage a property of any kind. It provides methods to retrieve and set the
/// property's identifier as well as its current value.\n
/// The interface was defined by Microsoft, but is not documented and never made it into any official
/// header file.
///
/// \sa ExplorerListView, IPropertyValueImpl
//////////////////////////////////////////////////////////////////////


#pragma once


// the interface's GUID
const IID IID_IPropertyValue = {0x7AF7F355, 0x1066, 0x4E17, {0xB1, 0xF2, 0x19, 0xFE, 0x2F, 0x09, 0x9C, 0xD2}};


class IPropertyValue :
    public IUnknown
{
public:
	/// \brief <em>Sets the \c PROPERTYKEY structure that identifies the property wrapped by the object</em>
	///
	/// \param[in] propKey The property identifier. It will be stored by the object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPropertyKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
	virtual HRESULT STDMETHODCALLTYPE SetPropertyKey(PROPERTYKEY const& propKey) = 0;
	/// \brief <em>Retrieves the \c PROPERTYKEY structure that identifies the property wrapped by the object</em>
	///
	/// \param[out] pPropKey Receives the property identifier.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetPropertyKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
	virtual HRESULT STDMETHODCALLTYPE GetPropertyKey(PROPERTYKEY* pPropKey) = 0;
	/// \brief <em>Retrieves the current value of the property wrapped by the object</em>
	///
	/// \param[out] pPropValue Receives the property value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa InitValue
	virtual HRESULT STDMETHODCALLTYPE GetValue(PROPVARIANT* pPropValue) = 0;
	/// \brief <em>Initializes the object with the property's current value</em>
	///
	/// \param[in] propValue The property's current value. It will be stored by the object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetValue
	virtual HRESULT STDMETHODCALLTYPE InitValue(PROPVARIANT const& propValue) = 0;
};