//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualListViewWorkAreasEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualListViewWorkAreasEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualListViewWorkAreas, ExLVwLibU::_IVirtualListViewWorkAreasEvents
/// \else
///   \sa VirtualListViewWorkAreas, ExLVwLibA::_IVirtualListViewWorkAreasEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualListViewWorkAreasEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualListViewWorkAreasEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualListViewWorkAreasEvents