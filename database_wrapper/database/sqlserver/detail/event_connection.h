#pragma once

#include "msado.h"

namespace msado
{
	template<typename Event>
	class connection_event
		: public ConnectionEventsVt
	{
	public:
		connection_event();
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

	private:
		ULONG ref_;
		Event event_;
	};
}

namespace msado
{
	template<typename Event>
	inline connection_event<Event>::connection_event()
		: ref_(0)
	{

	}

	//////////////////////////////////////////////////////
	//

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::QueryInterface(
		REFIID riid,
		void **ppvObject)
	{
		if (riid == IID_IUnknown || riid == __uuidof(ConnectionEventsVt))
		{
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		return E_FAIL;
	}

	template<typename Event>
	inline ULONG STDMETHODCALLTYPE connection_event<Event>::AddRef(void)
	{
		++ref_;
		return ref_;
	}

	template<typename Event>
	inline ULONG STDMETHODCALLTYPE connection_event<Event>::Release(void)
	{
		if (--ref_ == 0)
		{
			delete this;
		}
		return ref_;
	}

	//////////////////////////////////////////////////////
	//

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::InfoMessage(
		ADOError* pError, EventStatusEnum* adStatus, _ADOConnection* pConnection)
	{
		*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::BeginTransComplete(
		LONG TransactionLevel,
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADOConnection *pConnection)
	{
		event_.on_begin_transed(pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::CommitTransComplete(
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADOConnection *pConnection)
	{
		event_.on_commit_transed(pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::RollbackTransComplete(
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADOConnection *pConnection)
	{
		event_.on_rollback_transed(pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::WillExecute(
		BSTR *Source,
		CursorTypeEnum *CursorType,
		LockTypeEnum *LockType,
		long *Options,
		EventStatusEnum *adStatus,
		_ADOCommand *pCommand,
		_ADORecordset *pRecordset,
		_ADOConnection *pConnection)
	{
		event_.on_executing(*Source,
			static_cast<cursor_type>(*CursorType),
			static_cast<lock_type>(*LockType),
			*Options,
			pCommand,
			pRecordset,
			pConnection);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::ExecuteComplete(
		LONG RecordsAffected,
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADOCommand *pCommand,
		_ADORecordset *pRecordset,
		_ADOConnection *pConnection)
	{
		event_.on_executed(pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::WillConnect(
		BSTR *ConnectionString,
		BSTR *UserID,
		BSTR *Password,
		long *Options,
		EventStatusEnum *adStatus,
		_ADOConnection *pConnection)
	{
		event_.on_connecting(*ConnectionString, *UserID, *Password, *Options);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::ConnectComplete(
		ADOError *pError,
		EventStatusEnum *adStatus,
		_ADOConnection *pConnection)
	{
		event_.on_connected(pError);
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

	template<typename Event>
	inline HRESULT STDMETHODCALLTYPE connection_event<Event>::Disconnect(
		EventStatusEnum *adStatus,
		_ADOConnection *pConnection)
	{
		event_.on_disconnected();
		//*adStatus = adStatusUnwantedEvent;
		return S_OK;
	}

}
