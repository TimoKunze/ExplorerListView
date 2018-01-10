//////////////////////////////////////////////////////////////////////
/// \class Proxy_IListViewColumnEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IListViewColumnEvents interface</em>
///
/// \if UNICODE
///   \sa ListViewColumn, ExLVwLibU::_IListViewColumnEvents
/// \else
///   \sa ListViewColumn, ExLVwLibA::_IListViewColumnEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IListViewColumnEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IListViewColumnEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IListViewColumnEvents