//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualListViewColumnEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualListViewColumnEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualListViewColumn, ExLVwLibU::_IVirtualListViewColumnEvents
/// \else
///   \sa VirtualListViewColumn, ExLVwLibA::_IVirtualListViewColumnEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualListViewColumnEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualListViewColumnEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualListViewColumnEvents