//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewWorkAreaEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewWorkAreaEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewWorkArea, ExLVwLibU::_IListViewWorkAreaEvents
/// \else
///   \sa ListViewWorkArea, ExLVwLibA::_IListViewWorkAreaEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewWorkAreaEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewWorkAreaEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewWorkAreaEvents