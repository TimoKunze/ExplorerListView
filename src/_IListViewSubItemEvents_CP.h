//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewSubItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewSubItemEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewSubItem, ExLVwLibU::_IListViewSubItemEvents
/// \else
///   \sa ListViewSubItem, ExLVwLibA::_IListViewSubItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewSubItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewSubItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewSubItemEvents