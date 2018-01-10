//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualListViewItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualListViewItemEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualListViewItem, ExLVwLibU::_IVirtualListViewItemEvents
/// \else
///   \sa VirtualListViewItem, ExLVwLibA::_IVirtualListViewItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualListViewItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualListViewItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualListViewItemEvents