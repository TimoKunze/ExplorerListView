#ifndef DOXYGEN_SHOULD_SKIP_THIS
	#ifdef USE_STL
		#pragma once

		#include "stdafx.h"


		template<typename KeyDataType>
		struct ItemIndexHasher
		{
			enum
			{
				// increase this if you find the map getting too big
				bucket_size = 3,
				// make minimum size suitably big, to minimize rehashing
				min_buckets = 100000
			};

			size_t operator()(KeyDataType key) const
			{
				return static_cast<size_t>(key);
			}

			bool operator()(KeyDataType key1, KeyDataType key2) const
			{
				return (key1 < key2);
			}
		};
	#endif
#endif