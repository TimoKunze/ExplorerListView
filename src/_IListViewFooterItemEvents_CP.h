//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewFooterItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewFooterItemEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewFooterItem, ExLVwLibU::_IListViewFooterItemEvents
/// \else
///   \sa ListViewFooterItem, ExLVwLibA::_IListViewFooterItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewFooterItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewFooterItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewFooterItemEvents