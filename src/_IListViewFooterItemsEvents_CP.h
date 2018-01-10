//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewFooterItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewFooterItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewFooterItems, ExLVwLibU::_IListViewFooterItemsEvents
/// \else
///   \sa ListViewFooterItems, ExLVwLibA::_IListViewFooterItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewFooterItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewFooterItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewFooterItemsEvents