//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewGroupEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewGroupEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewGroup, ExLVwLibU::_IListViewGroupEvents
/// \else
///   \sa ListViewGroup, ExLVwLibA::_IListViewGroupEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewGroupEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewGroupEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewGroupEvents