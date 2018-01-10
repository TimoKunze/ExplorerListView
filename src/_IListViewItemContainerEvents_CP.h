//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewItemContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewItemContainerEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewItemContainer, ExLVwLibU::_IListViewItemContainerEvents
/// \else
///   \sa ListViewItemContainer, ExLVwLibA::_IListViewItemContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewItemContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewItemContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewItemContainerEvents