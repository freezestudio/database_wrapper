#pragma once

#include "msado.h"

namespace msado
{
	class recordset_event
		: public RecordsetEventsVt
	{
	public:
		explicit recordset_event(recordset* _event);
		recordset_event(recordset_event const&);
		~recordset_event();
	public:
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(
			/* [in] */ REFIID riid,
			/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
		virtual ULONG STDMETHODCALLTYPE AddRef(void);
		virtual ULONG STDMETHODCALLTYPE Release(void);
	public:
		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillChangeField(
			/* [in] */ LONG cFields,
			/* [in] */ VARIANT Fields,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);
		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FieldChangeComplete(
			/* [in] */ LONG cFields,
			/* [in] */ VARIANT Fields,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillChangeRecord(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ LONG cRecords,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RecordChangeComplete(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ LONG cRecords,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillChangeRecordset(
			/* [in] */ EventReasonEnum adReason,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RecordsetChangeComplete(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillMove(
			/* [in] */ EventReasonEnum adReason,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE MoveComplete(
			/* [in] */ EventReasonEnum adReason,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE EndOfRecordset(
			/* [out][in] */ __RPC__inout VARIANT_BOOL *fMoreData,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FetchProgress(
			/* [in] */ long Progress,
			/* [in] */ long MaxProgress,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FetchComplete(
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset);

	public:
		bool enable_event(bool enable = true);

	private:
		void init();
	private:
		ULONG ref_;
		DWORD dw_event_;
		recordset* eventp_;

	private:
		IConnectionPointContainer* pconnPointerContainer;
		IConnectionPoint* pconnPointer;
	};
}
