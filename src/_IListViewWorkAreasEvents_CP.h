//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewWorkAreasEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewWorkAreasEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewWorkAreas, ExLVwLibU::_IListViewWorkAreasEvents
/// \else
///   \sa ListViewWorkAreas, ExLVwLibA::_IListViewWorkAreasEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewWorkAreasEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewWorkAreasEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewWorkAreasEvents