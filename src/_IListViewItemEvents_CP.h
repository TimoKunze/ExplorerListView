//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewItemEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewItemEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewItem, ExLVwLibU::_IListViewItemEvents
/// \else
///   \sa ListViewItem, ExLVwLibA::_IListViewItemEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewItemEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewItemEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewItemEvents