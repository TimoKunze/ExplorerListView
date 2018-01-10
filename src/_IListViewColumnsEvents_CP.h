//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewColumnsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewColumnsEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewColumns, ExLVwLibU::_IListViewColumnsEvents
/// \else
///   \sa ListViewColumns, ExLVwLibA::_IListViewColumnsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewColumnsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewColumnsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewColumnsEvents