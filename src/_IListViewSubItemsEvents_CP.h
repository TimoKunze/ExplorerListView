//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewSubItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewSubItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewSubItems, ExLVwLibU::_IListViewSubItemsEvents
/// \else
///   \sa ListViewSubItems, ExLVwLibA::_IListViewSubItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewSubItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewSubItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewSubItemsEvents