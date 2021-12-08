#pragma once


class QuerySink : public IWbemObjectSink
{
    LONG m_lRef;
    bool bDone;
    CRITICAL_SECTION threadLock; // for thread safety

public:
    QuerySink() {
        m_lRef = 0; bDone = false;
        InitializeCriticalSection(&threadLock);
    }
    ~QuerySink() {
        bDone = true;
        DeleteCriticalSection(&threadLock);
    }

    virtual ULONG STDMETHODCALLTYPE AddRef();
    virtual ULONG STDMETHODCALLTYPE Release();
    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
                                                     void ** ppv);

    virtual HRESULT STDMETHODCALLTYPE Indicate(
        LONG lObjectCount,
        IWbemClassObject __RPC_FAR * __RPC_FAR * apObjArray
    );

    virtual HRESULT STDMETHODCALLTYPE SetStatus(
        /* [in] */ LONG lFlags,
        /* [in] */ HRESULT hResult,
        /* [in] */ BSTR strParam,
        /* [in] */ IWbemClassObject __RPC_FAR * pObjParam
    );

    bool IsDone();
};
