//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualListViewWorkAreaEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualListViewWorkAreaEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualListViewWorkArea, ExLVwLibU::_IVirtualListViewWorkAreaEvents
/// \else
///   \sa VirtualListViewWorkArea, ExLVwLibA::_IVirtualListViewWorkAreaEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualListViewWorkAreaEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualListViewWorkAreaEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualListViewWorkAreaEvents