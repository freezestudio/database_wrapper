#pragma once

namespace database
{
	//typedef struct tagSAFEARRAY {
	//	USHORT         cDims;
	//	USHORT         fFeatures;
	//	ULONG          cbElements;
	//	ULONG          cLocks;
	//	PVOID          pvData;
	//	SAFEARRAYBOUND rgsabound[1];
	//} SAFEARRAY, *LPSAFEARRAY;

	//typedef struct tagSAFEARRAYBOUND {
	//	ULONG cElements;
	//	LONG  lLbound;
	//} SAFEARRAYBOUND, *LPSAFEARRAYBOUND;

	//USHORT fFeatures values = one of follwing:
	//
	const static unsigned short feature_array_allocated_on_the_stack = FADF_AUTO;
	const static unsigned short feature_array_is_statically_allocated = FADF_STATIC;
	const static unsigned short feature_array_is_embedded_in_a_structure = FADF_EMBEDDED;
	const static unsigned short feature_array_may_not_be_resized_or_reallocated = FADF_FIXEDSIZE;
	//When set, there will be a pointer to the IRecordInfo interface at negative offset 4 in the array descriptor. 
	const static unsigned short feature_array_contains_records = FADF_RECORD;
	//When set, there will be a GUID at negative offset 16 in the safe array descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set. 
	const static unsigned short feature_array_has_an_iid_identifying_interface = FADF_HAVEIID;
	const static unsigned short feature_array_of_bstrs = FADF_BSTR;
	const static unsigned short feature_array_of_unknown_pointer = FADF_UNKNOWN;
	const static unsigned short feature_array_of_dispatch_pointer = FADF_DISPATCH;
	const static unsigned short feature_array_of_variants = FADF_VARIANT;
	const static unsigned short feature_array_bits_reserved_for_future_use = FADF_RESERVED;

	class safe_array
	{
	public:
		safe_array()
			: safe_arrayp_(nullptr)
			, boundp_(nullptr)
		{

		}

		explicit safe_array(SAFEARRAYBOUND* pbound)
			: safe_arrayp_(nullptr)
			, boundp_(pbound)
		{

		}
		
		operator SAFEARRAY*()
		{
			return safe_arrayp_;
		}

		SAFEARRAY* get()
		{
			return safe_arrayp_;
		}

		SAFEARRAY const* get() const
		{
			return safe_arrayp_;
		}

		SAFEARRAY** get_addressof()
		{
			return &safe_arrayp_;
		}
	
	public:
		//Increments the lock count of an array, 
		//and retrieves a pointer to the array data.
		HRESULT access_data(void* pData)
		{
			return ::SafeArrayAccessData(safe_arrayp_, &pData);
		}

		//Allocates memory for a safe array, 
		//based on a descriptor created with SafeArrayAllocDescriptor.
		HRESULT alloc_data()
		{
			return ::SafeArrayAllocData(safe_arrayp_);
		}

		//Allocates memory for a safe array descriptor.
		HRESULT alloc_descriptor(unsigned int dims)
		{
			return ::SafeArrayAllocDescriptor(dims, get_addressof());
		}

		//Creates a safe array descriptor for an array of any valid variant type,
		//including VT_RECORD, without allocating the array data.
		HRESULT alloc_descriptor_ex(unsigned short vt, unsigned int dims)
		{
			return ::SafeArrayAllocDescriptorEx(vt, dims, get_addressof());
		}

		//Creates a copy of an existing safe array.
		HRESULT copy(safe_array& arr)
		{
			return SafeArrayCopy(safe_arrayp_,arr.get_addressof());
		}

		//Copies the source array to the specified target array 
		//after releasing any resources in the target array. 
		//This is similar to SafeArrayCopy, 
		//except that the target array has to be set up by the caller. 
		//The target is not allocated or reallocated.
		HRESULT copy_data(safe_array& arr)
		{
			return ::SafeArrayCopyData(safe_arrayp_, (SAFEARRAY*)&arr);
		}

		//Creates a new array descriptor, 
		//allocates and initializes the data for the array, 
		//and returns a pointer to the new array descriptor.
		bool create(unsigned short vt,unsigned int dims)
		{
			safe_arrayp_=::SafeArrayCreate(vt, dims, boundp_);
			return safe_arrayp_ != nullptr;
		}

		bool create(unsigned short vt, unsigned int dims, SAFEARRAYBOUND* pbound)
		{
			boundp_ = pbound;
			safe_arrayp_ = ::SafeArrayCreate(vt, CDS_SET_PRIMARY, pbound);
			return safe_arrayp_ != nullptr;
		}

		//Creates and returns a safe array descriptor from the specified VARTYPE,
		//number of dimensions and bounds.
		bool create_ex(unsigned short vt, unsigned dims, void* pExtra)
		{
			safe_arrayp_ = ::SafeArrayCreateEx(vt, dims, boundp_, pExtra);
			return safe_arrayp_ != nullptr;
		}

		//Creates a one-dimensional array. A safe array created with SafeArrayCreateVector is a fixed size, so the constant FADF_FIXEDSIZE is always set.
		bool create_vector(unsigned short vt, long bound, unsigned long elements)
		{
			safe_arrayp_ = ::SafeArrayCreateVector(vt, bound, elements);
			return safe_arrayp_ != nullptr;
		}

		//Creates and returns a one-dimensional safe array of the specified VARTYPE and bounds.
		bool create_vector_ex(unsigned short vt, long bound, unsigned long elements, void* pExtra)
		{
			safe_arrayp_ = ::SafeArrayCreateVectorEx(vt, bound, elements, pExtra);
			return safe_arrayp_ != nullptr;
		}

		//Destroys an existing array descriptor and all of the data in the array. If objects are stored in the array, Release is called on each object in the array.
		HRESULT destroy()
		{
			return ::SafeArrayDestroy(safe_arrayp_);
		}

		//Destroys all the data in the specified safe array.
		HRESULT destroy_data()
		{
			return ::SafeArrayDestroyData(safe_arrayp_);
		}

		//Destroys the descriptor of the specified safe array.
		HRESULT destroy_descriptor()
		{
			return ::SafeArrayDestroyDescriptor(safe_arrayp_);
		}

		//Gets the number of dimensions in the array.
		unsigned get_dim() const
		{
			return ::SafeArrayGetDim(safe_arrayp_);
		}

		//Retrieves a single element of the array.
		HRESULT get_element(long& indices, void* pv)
		{
			return ::SafeArrayGetElement(safe_arrayp_, &indices, pv);
		}

		//Gets the size of an element.
		unsigned get_element_size() const
		{
			return ::SafeArrayGetElemsize(safe_arrayp_);
		}

		//Gets the GUID of the interface contained within the specified safe array.
		HRESULT get_iid(GUID& guid)
		{
			return ::SafeArrayGetIID(safe_arrayp_, &guid);
		}

		//Gets the lower bound for any dimension of the specified safe array.
		HRESULT get_l_bound(unsigned dim, long& bound)
		{
			return ::SafeArrayGetLBound(safe_arrayp_, dim, &bound);
		}

		//Retrieves the IRecordInfo interface of the UDT contained in the specified safe array.
		HRESULT get_record_info(IRecordInfo* pinfo)
		{
			return ::SafeArrayGetRecordInfo(safe_arrayp_, &pinfo);
		}

		//Gets the upper bound for any dimension of the specified safe array.
		HRESULT get_u_bound(unsigned dim, long& bound)
		{
			return ::SafeArrayGetUBound(safe_arrayp_, dim, &bound);
		}

		//Gets the VARTYPE stored in the specified safe array.
		HRESULT get_vartype(unsigned short& vt)
		{
			return ::SafeArrayGetVartype(safe_arrayp_, &vt);
		}

		//Increments the lock count of an array, 
		//and places a pointer to the array data in pvData of the array descriptor.
		HRESULT lock()
		{
			return ::SafeArrayLock(safe_arrayp_);
		}

		//Gets a pointer to an array element.
		HRESULT ptr_of_index(long& indices, void* pData)
		{
			return ::SafeArrayPtrOfIndex(safe_arrayp_, &indices, &pData);
		}

		//Stores the data element at the specified location in the array.
		HRESULT set_element(long& indices, void* pv)
		{
			return ::SafeArrayPutElement(safe_arrayp_,&indices, pv);
		}

		//Changes the right-most (least significant) bound of the specified safe array.
		HRESULT re_dim()
		{
			return ::SafeArrayRedim(safe_arrayp_, boundp_);
		}

		//Sets the GUID of the interface for the specified safe array.
		HRESULT set_iid(REFGUID guid)
		{
			return ::SafeArraySetIID(safe_arrayp_, guid);
		}

		//Sets the record info in the specified safe array.
		HRESULT set_record_info(IRecordInfo* pinfo)
		{
			return ::SafeArraySetRecordInfo(safe_arrayp_, pinfo);
		}

		//Decrements the lock count of an array, and invalidates the pointer retrieved by SafeArrayAccessData.
		HRESULT un_access_data()
		{
			return ::SafeArrayUnaccessData(safe_arrayp_);
		}

		//Decrements the lock count of an array so it can be freed or resized.
		HRESULT unlock()
		{
			return ::SafeArrayUnlock(safe_arrayp_);
		}
	private:
		SAFEARRAY* safe_arrayp_;
		SAFEARRAYBOUND* boundp_;
	};
}