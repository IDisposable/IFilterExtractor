// TextExtractor.h : Declaration of the CTextExtractor

#ifndef __TEXTEXTRACTOR_H_
#define __TEXTEXTRACTOR_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CTextExtractor
class ATL_NO_VTABLE CTextExtractor : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CTextExtractor, &CLSID_TextExtractor>,
	public ISupportErrorInfo,
	public IDelegatingDispImpl<ITextExtractor>
{
public:
	CTextExtractor()
    {
	}

DECLARE_REGISTRY_RESOURCEID(IDR_TEXTEXTRACTOR)
DECLARE_GET_CONTROLLING_UNKNOWN()

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CTextExtractor)
	COM_INTERFACE_ENTRY(ITextExtractor)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ITextExtractor
public:
	STDMETHOD(ExtractText)(/*[in]*/ BSTR fileName, /*[in]*/ long maxLength, /*[out, retval]*/ BSTR * fileText);
};

#endif //__TEXTEXTRACTOR_H_
