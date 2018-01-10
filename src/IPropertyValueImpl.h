//////////////////////////////////////////////////////////////////////
/// \class IPropertyValueImpl
/// \author Timo "TimoSoft" Kunze
/// \brief <em>An implementation of the \c IPropertyValue interface</em>
///
/// This class implements \c IPropertyValue and can be used to provide abstract access to a variant value.
/// The interface definition has been extracted from browseui.dll of Windows Vista, using an unknown
/// program (dumpbin?) which produces this output:
/// <pre>
/// Dump of file browseui.dll
///
/// File Type: DLL
///
/// BASE RELOCATIONS #6
///    D7000 RVA,       78 SizeOfBlock
///      5C0  DIR64      000007FF7A183200  ?QueryInterface@CPropertyValue@@UEAAJAEBU_GUID@@PEAPEAX@Z (public: virtual long __cdecl CPropertyValue::QueryInterface(struct _GUID const &,void * *))
///      5C8  DIR64      000007FF7A0B83DC  ?AddRef@CCancelWithEvent@@UEAAKXZ (public: virtual unsigned long __cdecl CCancelWithEvent::AddRef(void))
///      5D0  DIR64      000007FF7A183398  ?Release@CPropertyValue@@UEAAKXZ (public: virtual unsigned long __cdecl CPropertyValue::Release(void))
///      5D8  DIR64      000007FF7A183218  ?SetPropertyKey@CPropertyValue@@UEAAJAEBU_tagpropertykey@@@Z (public: virtual long __cdecl CPropertyValue::SetPropertyKey(struct _tagpropertykey const &))
///      5E0  DIR64      000007FF7A183238  ?GetPropertyKey@CPropertyValue@@UEAAJPEAU_tagpropertykey@@@Z (public: virtual long __cdecl CPropertyValue::GetPropertyKey(struct _tagpropertykey *))
///      5E8  DIR64      000007FF7A183260  ?GetValue@CPropertyValue@@UEAAJPEAUtagPROPVARIANT@@@Z (public: virtual long __cdecl CPropertyValue::GetValue(struct tagPROPVARIANT *))
///      5F0  DIR64      000007FF7A183278  ?InitValue@CPropertyValue@@UEAAJAEBUtagPROPVARIANT@@@Z (public: virtual long __cdecl CPropertyValue::InitValue(struct tagPROPVARIANT const &))
/// </pre>
/// The following WinDbg command should produce similar results: dt -v browseui!*.
/// However, public debug symbols do not include type information, so the quoted output must have been
/// created with private symbols. Also on Windows 7 CPropertyValue seems to be implemented in shell32.dll
/// instead of browseui.dll.
///
/// \todo Improve documentation of the \c pOuter parameter in \c CreateInstance().
/// \todo Improve documentation.
///
/// \sa ExplorerListView, IPropertyValue
//////////////////////////////////////////////////////////////////////

#pragma once

#include "IPropertyValue.h"


class IPropertyValueImpl :
    public IPropertyValue
{
public:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IUnknown
	///
	//@{
	/// \brief <em>Queries for an interface implementation</em>
	///
	/// Queries this object for its implementation of a given interface.
	///
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AddRef, Release,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682521.aspx">IUnknown::QueryInterface</a>
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Adds a reference to this object</em>
	///
	/// Increases the references counter of this object by 1.
	///
	/// \return The new reference count.
	///
	/// \sa Release, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms691379.aspx">IUnknown::AddRef</a>
	ULONG STDMETHODCALLTYPE AddRef();
	/// \brief <em>Removes a reference from this object</em>
	///
	/// Decreases the references counter of this object by 1. If the counter reaches 0, the object
	/// is destroyed.
	///
	/// \return The new reference count.
	///
	/// \sa AddRef, QueryInterface,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682317.aspx">IUnknown::Release</a>
	ULONG STDMETHODCALLTYPE Release();
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPropertyValue
	///
	//@{
	/// \brief <em>Sets the \c PROPERTYKEY structure that identifies the property wrapped by the object</em>
	///
	/// \param[in] propKey The property identifier. It will be stored by the object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPropertyKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
	HRESULT STDMETHODCALLTYPE SetPropertyKey(PROPERTYKEY const& propKey);
	/// \brief <em>Retrieves the \c PROPERTYKEY structure that identifies the property wrapped by the object</em>
	///
	/// \param[out] pPropKey Receives the property identifier.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetPropertyKey,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
	HRESULT STDMETHODCALLTYPE GetPropertyKey(PROPERTYKEY* pPropKey);
	/// \brief <em>Retrieves the current value of the property wrapped by the object</em>
	///
	/// \param[out] pPropValue Receives the property value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa InitValue
	HRESULT STDMETHODCALLTYPE GetValue(PROPVARIANT* pPropValue);
	/// \brief <em>Initializes the object with the property's current value</em>
	///
	/// \param[in] propValue The property's current value. It will be stored by the object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetValue
	HRESULT STDMETHODCALLTYPE InitValue(PROPVARIANT const& propValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Creates an instance of this class</em>
	///
	/// Creates an instance of this class, initializes the object with the given values and returns
	/// the implementation of a given interface.
	///
	/// \param[in] pOuter The outer unknown or controlling unknown of the aggregate.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppInterfaceImpl Receives the object's implementation of the interface identified
	///             by \c requiredInterface.
	///
	/// \return An \c HRESULT error code.
	static HRESULT CreateInstance(IUnknown* pOuter, REFIID requiredInterface, LPVOID* ppInterfaceImpl);
	/// \brief <em>Creates an instance of this class</em>
	///
	/// Creates an instance of this class, initializes the object with the given values and returns
	/// the \c IPropertyValue implementation.
	///
	/// \param[out] ppPropertyValue Receives the object's implementation of the \c IPropertyValue interface.
	///
	/// \return An \c HRESULT error code.
	static HRESULT CreateInstance(IPropertyValue** ppPropertyValue);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the object's reference count</em>
		volatile LONG referenceCount;
		/// \brief <em>Holds the \c PROPERTYKEY structure that identifies the wrapped property</em>
		///
		/// \sa SetPropertyKey, GetPropertyKey,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb773381.aspx">PROPERTYKEY</a>
		PROPERTYKEY propertyKey;
		/// \brief <em>Holds the current value of the wrapped property</em>
		///
		/// \sa InitValue, GetValue
		PROPVARIANT propertyValue;

		Properties()
		{
			referenceCount = 0;
			PropVariantInit(&propertyValue);
		}

		~Properties()
		{
			referenceCount = 0;
			PropVariantClear(&propertyValue);
		}
	} properties;
};     // IPropertyValueImpl