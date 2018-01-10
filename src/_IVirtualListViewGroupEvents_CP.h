//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualListViewGroupEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualListViewGroupEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualListViewGroup, ExLVwLibU::_IVirtualListViewGroupEvents
/// \else
///   \sa VirtualListViewGroup, ExLVwLibA::_IVirtualListViewGroupEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualListViewGroupEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualListViewGroupEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualListViewGroupEvents