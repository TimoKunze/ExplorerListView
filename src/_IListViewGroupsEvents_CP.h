//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewGroupsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewGroupsEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewGroups, ExLVwLibU::_IListViewGroupsEvents
/// \else
///   \sa ListViewGroups, ExLVwLibA::_IListViewGroupsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewGroupsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewGroupsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewGroupsEvents