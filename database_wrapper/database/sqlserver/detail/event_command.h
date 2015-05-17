#pragma once

#include "msado.h"

namespace msado
{
	class command_event
		: public ConnectionEventsVt
	{
	public:
		command_event(command* pevent);
		command_event(command_event const&);
		~command_event();
	public:
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(
			/* [in] */ REFIID riid,
			/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
		virtual ULONG STDMETHODCALLTYPE AddRef(void);
		virtual ULONG STDMETHODCALLTYPE Release(void);
	public:
		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE InfoMessage(
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE BeginTransComplete(
			/* [in] */ LONG TransactionLevel,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CommitTransComplete(
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RollbackTransComplete(
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillExecute(
			/* [out][in] */ __RPC__deref_inout_opt BSTR *Source,
			/* [out][in] */ __RPC__inout CursorTypeEnum *CursorType,
			/* [out][in] */ __RPC__inout LockTypeEnum *LockType,
			/* [out][in] */ __RPC__inout long *Options,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOCommand *pCommand,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ExecuteComplete(
			/* [in] */ LONG RecordsAffected,
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOCommand *pCommand,
			/* [in] */ __RPC__in_opt _ADORecordset *pRecordset,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WillConnect(
			/* [out][in] */ __RPC__deref_inout_opt BSTR *ConnectionString,
			/* [out][in] */ __RPC__deref_inout_opt BSTR *UserID,
			/* [out][in] */ __RPC__deref_inout_opt BSTR *Password,
			/* [out][in] */ __RPC__inout long *Options,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ConnectComplete(
			/* [in] */ __RPC__in_opt ADOError *pError,
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

		virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Disconnect(
			/* [out][in] */ __RPC__inout EventStatusEnum *adStatus,
			/* [in] */ __RPC__in_opt _ADOConnection *pConnection);

	public:
		bool enable_event(bool enable = true);

	private:
		void init();
	private:
		ULONG ref_;
		DWORD dw_event_;
		command* eventp_;

	private:
		IConnectionPointContainer* pconnPointerContainer;
		IConnectionPoint* pconnPointer;
	};
}
