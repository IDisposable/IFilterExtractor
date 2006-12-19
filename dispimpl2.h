// dispimpl2.h: Alternative implementations of IDispatch
/////////////////////////////////////////////////////////////////////////////
// Copyright (c) 1998-2003 Chris Sells
// All rights reserved.
/////////////////////////////////////////////////////////////////////////////
// History:
// 6/11/03:
//  -Ehab Kashkash (ekashkash@hotmail.com): Added "typename" keyword 
//   when initializing static variables of user-defined types (_tihclass)
//   in order to be able to compile with VS.NET.
//
// 3/3/00:
//  -Added some description of using IDelegatingDispImpl with controls to
//   satisfied picky control contains (like VB). Thanks to Greg Kedge
//   [Greg.Kedge@usa.xerox.com] for providing the initial draft.
//
// 7/22/99 (later again...):
//  -Simon Fell (atl@ZAKS.DEMON.CO.UK) pointed out a bug where
//   I don't actually derived from D in IDelegatingDispImpl. Doh!
//
// 7/22/99 (sometime later):
//  -Put optional D back in as a base, but default to IDispatch,
//   allowing you to use IDelegatingDispImpl like so:
//
//     class CFoo :
//         public IDelegatingDispImpl<IFoo, &IID_IFoo, DFoo>...
//
// 7/22/99:
//  -Simplified usage of IDelegatingDispImpl. You will
//   have to change your usage from:
//
//     class CFoo :
//         public IDelegatingDispImpl<DFoo, IFoo>...
//
//   to:
//
//     class CFoo :
//         public IDelegatingDispImpl<IFoo>...
//
// 10/25/98:
//  -Initial release.
//
// NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK.
//
// Contact the author with suggestions or comments at csells@sellsbrothers.com.
//
/////////////////////////////////////////////////////////////////////////////
//
// This header defines three C++ classes for use in implementing
// dispatch-based interfaces:
//
//  -IDualDispImpl for implementing dual interfaces.
//
//  -IDelegatingDispImpl for implementing IDispatch by delegation
//   to another interface (typically a custom interface).
//
//  -IRawDispImpl for implementing raw dispinterfaces.
//
// These classes are useful because ATL's IDispatchImpl can
// only implement duals.
//
/////////////////////////////////////////////////////////////////////////////
//
// IDualDispImpl: For use with dispatch-based interfaces declared like so:
//
// [dual]
// interface IFoo : IDispatch
// {
//    ...
// }
//
// IDualDispImpl implements all four IDispatch methods.
// IDualDispImpl gets the IDispatch vtbl entries by deriving from
// the IDispatch that servers as the base class for the dual interface.
// IDualDispImpl is just like ATL's IDispatchImpl (in fact, it
// derives from IDispatchImpl), but it uses __uuidof to make the usage
// simplier.
//
// Usage:
//  class CFoo : ..., public IDualDispImpl<IFoo>
//
/////////////////////////////////////////////////////////////////////////////
//
// IDelegatingDispImpl: For implementing IDispatch in terms of another
// (typically custom) interface, e.g.:
//
// [oleautomation]
// interface IFoo : IUnknown
// {
//    ...
// }
//
// IDelegatingDispImpl implements all four IDispatch methods.
// IDelegatingDispImpl gets the IDispatch vtbl entries by deriving from
// IDispatch in addition to the implementation interface.
//
// Usage:
//  class CFoo : ..., public IDelegatingDispImpl<IFoo>
//
// In the case where the coclass is intended to represent a control,
// there is a need for the coclass to have a [default] dispinterface.
// Otherwise, some control containers (notably VB) throw arcane error when
// the control is loaded.  For a control that you intend to provide the
// custom interface and delegating dispatch mechanism, you will have to
// provide a dispinterface defined in terms of a custom interface like
// so:
//
// dispinterface DFoo
// {
//    interface IFoo;
// }
//
// coclass Foo
// {
//  [default] interface DFoo;
//  interface IFoo;
// };
//
// For every other situation, declaring a dispinterface in terms of a
// custom interface is not necessary to use IDelegatingDispatchImpl.
// However, if you'd like DFoo to be in the base class list (as needed
// for the caveat control), you may use DFoo as the base class instead
// of the default template argument IDispatch like so:
//
// Usage:
//  class CFoo : ..., public IDelegatingDispImpl<IFoo, &IID_IFoo, DFoo>
//
/////////////////////////////////////////////////////////////////////////////
//
// IRawDispImpl: For use with dispatch-based interfaces declared like so:
//
// dispinterface DFoo
// {
// properties:
//   ...
// methods:
//   ...
// }
//
// IRawDispImpl implements three of the IDispatch methods.
// Invoke is left to the implementor.
//
// Usage:
//  class CFoo : ..., public IRawDispImpl<DFoo>
//  {
//  ...
//      STDMETHODIMP Invoke(...);   // Implemented by hand.
//  };
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#ifndef INC_DISPIMPL2
#define INC_DISPIMPL2

/////////////////////////////////////////////////////////////////////////////
// IDualDispImpl

template <class T, const IID* piid = &__uuidof(T), const GUID* plibid = &CComModule::m_libid, WORD wMajor = 1,
WORD wMinor = 0, class tihclass = CComTypeInfoHolder>
class ATL_NO_VTABLE IDualDispImpl : public IDispatchImpl<T, piid, plibid, wMajor, wMinor, tihclass> {};

/////////////////////////////////////////////////////////////////////////////
// IDelegatingDispImpl

template <class T, const IID* piid = &__uuidof(T), class D = IDispatch,
const GUID* plibid = &CComModule::m_libid, WORD wMajor = 1,
WORD wMinor = 0, class tihclass = CComTypeInfoHolder>
class ATL_NO_VTABLE IDelegatingDispImpl : public T, public D
{
public:
    typedef tihclass _tihclass;

    // IDispatch
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo)
    {
        *pctinfo = 1;
        return S_OK;
    }

    STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
    {
        return _tih.GetTypeInfo(itinfo, lcid, pptinfo);
    }

    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
        LCID lcid, DISPID* rgdispid)
    {
        return _tih.GetIDsOfNames(riid, rgszNames, cNames, lcid, rgdispid);
    }

    STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid,
        LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo, UINT* puArgErr)
    {
        // NOTE: reinterpret_cast because CComTypeInfoHolder makes the mistaken
        //       assumption that the typeinfo can only Invoke using an IDispatch*.
        //       Since the implementation only passes the itf onto
        //       ITypeInfo::Invoke (which takes a void*), this is a safe cast
        //       until the ATL team fixes CComTypeInfoHolder.
        return _tih.Invoke(reinterpret_cast<IDispatch*>(static_cast<T*>(this)),
            dispidMember, riid, lcid, wFlags, pdispparams,
            pvarResult, pexcepinfo, puArgErr);
    }

protected:
    static _tihclass _tih;

    static HRESULT GetTI(LCID lcid, ITypeInfo** ppInfo)
    {
        return _tih.GetTI(lcid, ppInfo);
    }
};

template <class T, const IID* piid, class D, const GUID* plibid, WORD wMajor, WORD wMinor, class tihclass>
typename IDelegatingDispImpl<T, piid, D, plibid, wMajor, wMinor, tihclass>::_tihclass
IDelegatingDispImpl<T, piid, D, plibid, wMajor, wMinor, tihclass>::_tih =
{ piid, plibid, wMajor, wMinor, NULL, 0, NULL, 0 };

/////////////////////////////////////////////////////////////////////////////
// IRawDispImpl

template <class T, const IID* piid = &__uuidof(T), const GUID* plibid = &CComModule::m_libid, WORD wMajor = 1,
WORD wMinor = 0, class tihclass = CComTypeInfoHolder>
class ATL_NO_VTABLE IRawDispImpl : public T
{
public:
    typedef tihclass _tihclass;

    // IDispatch
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo)
    {
        *pctinfo = 1;
        return S_OK;
    }

    STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
    {
        return _tih.GetTypeInfo(itinfo, lcid, pptinfo);
    }

    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
        LCID lcid, DISPID* rgdispid)
    {
        return _tih.GetIDsOfNames(riid, rgszNames, cNames, lcid, rgdispid);
    }

    // You have to implement Invoke by hand.
    // However, the same techniques used for IDispEventImpl
    // could be used to implement this as well.
    STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr) = 0;

protected:
    static _tihclass _tih;

    static HRESULT GetTI(LCID lcid, ITypeInfo** ppInfo)
    {
        return _tih.GetTI(lcid, ppInfo);
    }
};

template <class T, const IID* piid, const GUID* plibid, WORD wMajor, WORD wMinor, class tihclass>
typename IRawDispImpl<T, piid, plibid, wMajor, wMinor, tihclass>::_tihclass
IRawDispImpl<T, piid, plibid, wMajor, wMinor, tihclass>::_tih =
{piid, plibid, wMajor, wMinor, NULL, 0, NULL, 0};

#endif  // INC_DISPIMPL2
