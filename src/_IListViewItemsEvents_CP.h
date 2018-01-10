//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewItemsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewItemsEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewItems, ExLVwLibU::_IListViewItemsEvents
/// \else
///   \sa ListViewItems, ExLVwLibA::_IListViewItemsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewItemsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewItemsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewItemsEvents