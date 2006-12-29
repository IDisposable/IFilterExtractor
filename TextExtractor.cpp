// TextExtractor.cpp : Implementation of CTextExtractor
#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

#include <string>
#include <sstream>

#include "dispimpl2.h"
#include "ExtractText.h"
#include "TextExtractor.h"
#include "Filter.h"
#include "FiltErr.h"
#include "NTQuery.h"

/////////////////////////////////////////////////////////////////////////////
// CTextExtractor

STDMETHODIMP CTextExtractor::InterfaceSupportsErrorInfo(REFIID riid)
{
   static const IID* arr[] =
   {
      &IID_ITextExtractor
   };
   for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
   {
      if (InlineIsEqualGUID(*arr[i],riid))
         return S_OK;
   }
   return S_FALSE;
}

// Per W3C spec http://www.w3.org/TR/REC-xml#charsets
// Valid characters are #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] |
//                      [#x10000-#x10FFFF]

inline static void ValidUnicode(wchar_t & ch)
{
   if (ch < 0x0020)     // if less than ASCII space
   {
      if ((ch == 0x000D)      // CR
         || (ch == 0x000A)    // or LF
         || (ch == 0x0009))   // or TAB
         return;                 // it's valid!
      else
         ch = L' ';              // morph to blank
   }
   else if (ch > 0x007e) // or greater than ASCII '~' 
   {
      if (ch <= 0xD7FF)
         return;                 // it's valid!
      else if (ch >= 0xF8FF && ch <= 0xFFFD)
         return;                 // it's valid!
      else
         ch = L' ';              // morph to blank
      
      // note we don't support surrogates, private use or high-Unicode 0x10000-0x10FFFF characters
   }
   else
      return;                    // it's valid!
}

static void CleanUpCharacters(size_t chBuf, wchar_t *buf)
{
   // The game here is to fold any "cute" versions of characters to thier 
   // simplified form to make parsing easier.

   buf[chBuf] = 0;   // must be null terminated..

   for (size_t i = 0; i < chBuf; ++i)
   {
      wchar_t & ch = buf[i];

      switch (ch)
      {
         case 0:        // embedded null
         case 0x2000:   // en quad
         case 0x2001:   // em quad
         case 0x2002:   // en space
         case 0x2003:   // em space
         case 0x2004:   // three-per-em space
         case 0x2005:   // four-per-em space
         case 0x2006:   // six-per-em space
         case 0x2007:   // figure space
         case 0x2008:   // puctuation space
         case 0x2009:   // thin space
         case 0x200A:   // hair space
         case 0x200B:   // zero-width space
         case 0x200C:   // zero-width non-joiner
         case 0x200D:   // zero-width joiner
         case 0x202f:   // no-break space
         case 0x3000:   // ideographic space
            ch = L' ';
            break;

         case 0x00B6:   // pilcro
         case 0x2028:   // line seperator
         case 0x2029:   // paragraph seperator
            ch = L'\n';
            break;

         case 0x00AD:   // soft-hyphen
         case 0x00B7:   // middle dot
         case 0x2010:   // hyphen
         case 0x2011:   // non-breaking hyphen
         case 0x2012:   // figure dash
         case 0x2013:   // en dash
         case 0x2014:   // em dash
         case 0x2015:   // quote dash
         case 0x2027:   // hyphenation point
         case 0x2043:   // hyphen bullet
         case 0x208B:   // subscript minus
         case 0xFE31:   // vertical em dash
         case 0xFE32:   // vertical en dash
         case 0xFE58:   // small em dash
         case 0xFE63:   // small hyphen minus
            ch = L'-';
            break;

         case 0x00B0:   // degree
         case 0x2018:   // left single quote
         case 0x2019:   // right single quote
         case 0x201A:   // low right single quote
         case 0x201B:   // high left single quote
         case 0x2032:   // prime
         case 0x2035:   // reversed prime
         case 0x2039:   // left-pointing angle quotation mark
         case 0x203A:   // right-pointing angle quotation mark
            ch = L'\'';
            break;
            
         case 0x201C:   // left double quote
         case 0x201D:   // right double quote
         case 0x201E:   // low right double quote
         case 0x201F:   // high left double quote
         case 0x2033:   // double prime
         case 0x2034:   // triple prime
         case 0x2036:   // reversed double prime
         case 0x2037:   // reversed triple prime
         case 0x00AB:   // left-pointing double angle quotation mark
         case 0x00BB:   // right-pointing double angle quotation mark
         case 0x3003:   // ditto mark
         case 0x301D:   // reversed double prime quotation mark
         case 0x301E:   // double prime quotation mark
         case 0x301F:   // low double prime quotation mark
            ch = L'\"';
            break;
            
         case 0x00A7:   // section-sign
         case 0x2020:   // dagger
         case 0x2021:   // double-dagger
         case 0x2022:   // bullet
         case 0x2023:   // triangle bullet
         case 0x203B:   // reference mark
         case 0xFE55:   // small colon
            ch = L':';
            break;

         case 0x2024:   // one dot leader
         case 0x2025:   // two dot leader
         case 0x2026:   // elipsis
         case 0x3002:   // ideographic full stop
         case 0xFE30:   // two dot vertical leader
         case 0xFE52:   // small full stop
            ch = L'.';
            break;

         case 0x3001:   // ideographic comma
         case 0xFE50:   // small comma
         case 0xFE51:   // small ideographic comma
            ch = L',';
            break;
            
         case 0xFE54:   // small semicolon
            ch = L';';
            break;

         case 0x00A6:   // broken-bar
         case 0x2016:   // double vertical line
            ch = L'|';
            break;

         case 0x2017:   // double low line
         case 0x203E:   // overline
         case 0x203F:   // undertie
         case 0x2040:   // character tie
         case 0xFE33:   // vertical low line
         case 0xFE49:   // dashed overline
         case 0xFE4A:   // centerline overline
         case 0xFE4D:   // dashed low line
         case 0xFE4E:   // centerline low line
            ch = L'_';
            break;
            
         case 0x301C:   // wave dash
         case 0x3030:   // wavy dash
         case 0xFE34:   // vertical wavy low line
         case 0xFE4B:   // wavy overline
         case 0xFE4C:   // double wavy overline
         case 0xFE4F:   // wavy low line
            ch = L'~';
            break;
            
         case 0x2038:   // caret
         case 0x2041:   // caret insertion point
            ch = L'^';
            break;

         case 0x2030:   // per-mille
         case 0x2031:   // per-ten thousand
         case 0xFE6A:   // small per-cent
            ch = L'%';
            break;
            
         case 0xFE6B:   // small commercial at
            ch = L'@';
            break;
            
         case 0x00A9:   // copyright
            ch = L'c';
            break;

         case 0x00B5:   // micro
            ch = L'u';
            break;
   
         case 0x00AE:   // registered
            ch = L'r';
            break;

         case 0x207A:   // superscript plus
         case 0x208A:   // subscript plus
         case 0xFE62:   // small plus
            ch = L'+';
            break;
            
         case 0x2044:   // fraction slash
            ch = L'/';
            break;

         case 0x2042:   // asterism
         case 0xFE61:   // small asterisk
            ch = L'*';
            break;
            
         case 0x208C:   // subscript equal
         case 0xFE66:   // small equal
            ch = L'=';
            break;
            
         case 0xFE68:   // small reverse solidus
            ch = L'\\';
            break;
            
         case 0xFE5F:   // small number sign
            ch = L'#';
            break;
            
         case 0xFE60:   // small ampersand
            ch = L'&';
            break;
            
         case 0xFE69:   // small dollar sign
            ch = L'$';
            break;
            
         case 0x2045:   // left square bracket with quill
         case 0x3010:   // left black lenticular bracket
         case 0x3016:   // left white lenticular bracket
         case 0x301A:   // left white square bracket
         case 0xFE3B:   // vertical left lenticular bracket
         case 0xFF41:   // vertical left corner bracket
         case 0xFF43:   // vertical white left corner bracket
            ch = L'[';
            break;
            
         case 0x2046:   // right square bracket with quill
         case 0x3011:   // right black lenticular bracket
         case 0x3017:   // right white lenticular bracket
         case 0x301B:   // right white square bracket
         case 0xFE3C:   // vertical right lenticular bracket
         case 0xFF42:   // vertical right corner bracket
         case 0xFF44:   // vertical white right corner bracket
            ch = L']';
            break;
            
         case 0x208D:   // subscript left parenthesis
         case 0x3014:   // left tortise-shell bracket
         case 0x3018:   // left white tortise-shell bracket
         case 0xFE35:   // vertical left parenthesis
         case 0xFE39:   // vertical left tortise-shell bracket
         case 0xFE59:   // small left parenthesis
         case 0xFE5D:   // small left tortise-shell bracket
            ch = L'(';
            break;
            
         case 0x208E:   // subscript right parenthesis
         case 0x3015:   // right tortise-shell bracket
         case 0x3019:   // right white tortise-shell bracket
         case 0xFE36:   // vertical right parenthesis
         case 0xFE3A:   // vertical right tortise-shell bracket
         case 0xFE5A:   // small right parenthesis
         case 0xFE5E:   // small right tortise-shell bracket
            ch = L')';
            break;
            
         case 0x3008:   // left angle bracket
         case 0x300A:   // left double angle bracket
         case 0xFF3D:   // vertical left double angle bracket
         case 0xFF3F:   // vertical left angle bracket
         case 0xFF64:   // small less-than
            ch = L'<';
            break;
            
         case 0x3009:   // right angle bracket
         case 0x300B:   // right double angle bracket
         case 0xFF3E:   // vertical right double angle bracket
         case 0xFF40:   // vertical right angle bracket
         case 0xFF65:   // small greater-than
            ch = L'>';
            break;
            
         case 0xFE37:   // vertical left curly bracket
         case 0xFE5B:   // small left curly bracket
            ch = L'{';
            break;
            
         case 0xFE38:   // vertical right curly bracket
         case 0xFE5C:   // small right curly bracket
            ch = L'}';
            break;
            
         case 0x00A1:   // inverted exclamation mark
         case 0x00AC:   // not
         case 0x203C:   // double exclamation mark
         case 0x203D:   // interrobang
         case 0xFE57:   // small exclamation mark
            ch = L'!';
            break;

         case 0x00BF:   // inverted question mark
         case 0xFE56:   // small question mark
            ch = L'?';
            break;

         case 0x00B9:   // superscript one
            ch = L'1';
            break;

         case 0x00B2:   // superscript two
            ch = L'2';
            break;
            
         case 0x00B3:   // superscript three
            ch = L'3';
            break;

         case 0x2070:   // superscript zero
         case 0x2074:   // superscript four
         case 0x2075:   // superscript five
         case 0x2076:   // superscript six
         case 0x2077:   // superscript seven
         case 0x2078:   // superscript eight
         case 0x2079:   // superscript nine
         case 0x2080:   // subscript zero
         case 0x2081:   // subscript one
         case 0x2082:   // subscript two
         case 0x2083:   // subscript three
         case 0x2084:   // subscript four
         case 0x2085:   // subscript five
         case 0x2086:   // subscript six
         case 0x2087:   // subscript seven
         case 0x2088:   // subscript eight
         case 0x2089:   // subscript nine
         case 0x3021:   // Hangzhou numeral one
         case 0x3022:   // Hangzhou numeral two
         case 0x3023:   // Hangzhou numeral three
         case 0x3024:   // Hangzhou numeral four
         case 0x3025:   // Hangzhou numeral five
         case 0x3026:   // Hangzhou numeral six
         case 0x3027:   // Hangzhou numeral seven
         case 0x3028:   // Hangzhou numeral eight
         case 0x3029:   // Hangzhou numeral nine
            ch = (ch & 0x000F) + L'0';
            break;

         // ONE is at ZERO location... careful
         case 0x3220:   // parenthesized ideograph one
         case 0x3221:   // parenthesized ideograph two
         case 0x3222:   // parenthesized ideograph three
         case 0x3223:   // parenthesized ideograph four
         case 0x3224:   // parenthesized ideograph five
         case 0x3225:   // parenthesized ideograph six
         case 0x3226:   // parenthesized ideograph seven
         case 0x3227:   // parenthesized ideograph eight
         case 0x3228:   // parenthesized ideograph nine
         case 0x3280:   // circled ideograph one
         case 0x3281:   // circled ideograph two
         case 0x3282:   // circled ideograph three
         case 0x3283:   // circled ideograph four
         case 0x3284:   // circled ideograph five
         case 0x3285:   // circled ideograph six
         case 0x3286:   // circled ideograph seven
         case 0x3287:   // circled ideograph eight
         case 0x3288:   // circled ideograph nine
            ch = (ch & 0x000F) + L'1';
            break;
            
         case 0x3007:   // ideographic number zero
         case 0x24EA:   // circled number zero
            ch = L'0';
            break;
            
         default:
            if (0xFF01 <= ch           // fullwidth exclamation mark 
                && ch <= 0xFF5E)       // fullwidth tilde
            {
               // the fullwidths line up with ASCII low subset
               ch = ch & 0xFF00 + L'!' - 1;               
            }
            else if (0x2460 <= ch      // circled one
                     && ch <= 0x2468)  // circled nine
            {
               ch = ch - 0x2460 + L'1';
            }
            else if (0x2474 <= ch      // parenthesized one
                     && ch <= 0x247C)  // parenthesized nine
            {
               ch = ch - 0x2474 + L'1';
            }
            else if (0x2488 <= ch      // one full stop
                     && ch <= 0x2490)  // nine full stop
            {
               ch = ch - 0x2488 + L'1';
            }
            else if (0x249C <= ch      // parenthesized small a
                     && ch <= 0x24B5)  // parenthesized small z
            {
               ch = ch - 0x249C + L'a';
            }
            else if (0x24B6 <= ch      // circled capital A
                     && ch <= 0x24CF)  // circled capital Z
            {
               ch = ch - 0x24B6 + L'A';
            }
            else if (0x24D0 <= ch      // circled small a
                     && ch <= 0x24E9)  // circled small z
            {
               ch = ch - 0x24D0 + L'a';
            }
            else if (0x2500 <= ch      // box drawing (begin)
                     && ch <= 0x257F)  // box drawing (end)
            {
               ch = L'|';
            }
            else if (0x2580 <= ch      // block elements (begin)
                     && ch <= 0x259F)  // block elements (end)
            {
               ch = L'#';
            }
            else if (0x25A0 <= ch      // geometric shapes (begin)
                     && ch <= 0x25FF)  // geometric shapes (end)
            {
               ch = L'*';
            }
            else if (0x2600 <= ch      // dingbats (begin)
                     && ch <= 0x267F)  // dingbats (end)
            {
               ch = L'.';
            }
            else
               ValidUnicode(ch);   // validate that it's legit Unicode
            break;
      }
   }
}

STDMETHODIMP CTextExtractor::ExtractText(BSTR fileName, long maxLength, BSTR * fileText)
{
   if (NULL == fileName)
      return E_POINTER;

   if (0 == ::SysStringLen(fileName))
      return E_INVALIDARG;

   if (NULL == fileText)
      return E_POINTER;

   *fileText = NULL;

   HRESULT hr = E_UNEXPECTED;

   try
   {
      CComPtr<IUnknown> spIUnk;
      hr = LoadIFilter(fileName, NULL, reinterpret_cast<void**>(&spIUnk));

      if (SUCCEEDED(hr))
      {
         CComQIPtr<IFilter> spIFilter = spIUnk;

         if (spIFilter)
         {
            DWORD dwFlags = 0;
            hr = spIFilter->Init(IFILTER_INIT_CANON_PARAGRAPHS |
                                 IFILTER_INIT_CANON_HYPHENS |
                                 IFILTER_INIT_CANON_SPACES |
                                 IFILTER_INIT_APPLY_INDEX_ATTRIBUTES |
                                 IFILTER_INIT_INDEXING_ONLY,
                                 0, NULL, &dwFlags);
            AtlTrace(_T("Init() hr=%x, dwFlags=%x\n"), hr, dwFlags);

            if (SUCCEEDED(hr))
            {
               std::wstringstream wout;

               STAT_CHUNK statChunk;
               memset(&statChunk, 0, sizeof(statChunk));

               bool moreChunks = true;

               while (moreChunks)
               {
                  if (maxLength && wout.tellp() > maxLength)
                  {
                     hr = S_FALSE;
                     break;
                  }

                  hr = spIFilter->GetChunk(&statChunk);
                  AtlTrace(_T("GetChunk() hr=%x, breakType=%d, flags=%x\n"), hr, statChunk.breakType, statChunk.flags);

                  if (SUCCEEDED(hr))
                  {
                     // ignore non-text chunks...
                     if (CHUNK_TEXT != (CHUNK_TEXT & statChunk.flags))
                        continue;

                     switch (statChunk.breakType)
                     {
                        case CHUNK_NO_BREAK:
                           break;

                        case CHUNK_EOW:
                           wout << L' ';
                           break;

                        case CHUNK_EOS:
                        case CHUNK_EOC:
                        case CHUNK_EOP:
                           wout << L"\r\n";
                           break;
                     }

                     bool moreText = true;

                     while (moreText)
                     {
                        static const int cChunkSize = 4096;

                        wchar_t buf[cChunkSize + 1];
                        unsigned long chBuf = cChunkSize;
                        memset(buf, 0, sizeof(buf));

                        hr = spIFilter->GetText(&chBuf, buf);
                        AtlTrace(_T("GetText() hr=%x, chBuf=%d\n"), hr, chBuf);

                        if (SUCCEEDED(hr))
                        {
                           CleanUpCharacters(chBuf, buf);
                           wout << buf;

                           if (maxLength && wout.tellp() > maxLength)
                           {
                              hr = S_FALSE;
                              moreText = false;
                           }

                           if (FILTER_S_LAST_TEXT == hr)
                           {
                              hr = S_OK;
                              moreText = false;
                           }
                        }
                        else
                        {
                           switch (hr)
                           {
                              case FILTER_E_NO_MORE_TEXT:
                                 hr = S_OK;
                                 moreText = false;
                                 break;
                              
                              case FILTER_E_NO_TEXT:
                                 return Error("GetText: The current chunk does not contain text.", __uuidof(TextExtractor), hr);

                              default:
                                 return Error("GetText: Unexpected error.", __uuidof(TextExtractor), hr);
                           }
                        }
                     }
                  }
                  else
                  {
                     switch (hr)
                     {
                        case FILTER_E_EMBEDDING_UNAVAILABLE:
                        case FILTER_E_LINK_UNAVAILABLE:
                           continue;   // next chunk...

                        case FILTER_E_END_OF_CHUNKS:
                           hr = S_OK;
                           moreChunks = false;
                           continue;

                        case FILTER_E_PASSWORD:
                           return Error("GetChunk: Password or other security-related access failure.", __uuidof(TextExtractor), hr);

                        case FILTER_E_ACCESS:
                           return Error("GetChunk: Access failure.", __uuidof(TextExtractor), hr);

                        default:
                           return Error("GetChunk: Unexpected error.", __uuidof(TextExtractor), hr);
                     }
                  }
               }

               *fileText = SysAllocString(wout.str().c_str());
            }
            else
            {
               switch (hr)
               {
                  case E_FAIL:
                     return Error("Init: File filter for this file type not found.", __uuidof(TextExtractor), hr);

                  case E_INVALIDARG:
                     return Error("Init: Count and contents of attributes do not agree.", __uuidof(TextExtractor), hr);

                  case FILTER_E_PASSWORD:
                     return Error("Init: Access has been denied because of password protection or similar security measures.", __uuidof(TextExtractor), hr);

                  case FILTER_E_ACCESS:
                     return Error("Init: Unable to access file.", __uuidof(TextExtractor), hr);

                  default:
                     return Error("Init: Unexpected error.", __uuidof(TextExtractor), hr);
               }
            }
         }
         else
         {
            return Error("Unable to query for IFilter interface.", __uuidof(TextExtractor), E_FAIL);
         }
      }
      else
      {
         switch (hr)
         {
            case E_ACCESSDENIED:
               return Error("LoadIFilter: Access denied to the filter file.", __uuidof(TextExtractor), hr);

            case E_HANDLE:
               return Error("LoadIFilter: Invalid handle, probably due to a low-memory situation.", __uuidof(TextExtractor), hr);

            case E_INVALIDARG:
               return Error("LoadIFilter: Invalid parameter.", __uuidof(TextExtractor), hr);

            case E_OUTOFMEMORY:
               return Error("LoadIFilter: Insufficient memory or other resources to complete the operation.", __uuidof(TextExtractor), hr);

            case E_FAIL:
               return Error("LoadIFilter: Unknown error.", __uuidof(TextExtractor), hr);
               
            case FILTER_E_PASSWORD:
               return Error("LoadIFilter: Access has been denied because of password protection or similar security measures.", __uuidof(TextExtractor), hr);
               
            case FILTER_E_ACCESS:
               return Error("LoadIFilter: Unable to access file.", __uuidof(TextExtractor), hr);               

            default:
               return Error("LoadIFilter: Unexpected error.", __uuidof(TextExtractor), hr);
         }
      }
   }
   catch (...)
   {
      return Error("Unexpected exception",  __uuidof(TextExtractor), E_FAIL);
   }

   return hr;
}
