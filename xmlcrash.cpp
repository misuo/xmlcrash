// cl / EHsc / Zi / D DEBUG xmlcrash.cpp

// C++ includes
#include <string>   // For access to std::string
#include <fstream>  // For access to std::ifstream
#include <iostream> // For access to std::cout
#include <sstream>  // For access to std::ostringstream
#include <vector>   // For access to std::vector
#include <cassert>  // For access to assert().

// Windows includes
#include <windows.h>
#include <objidl.h>  // For access to IStream class

// Type library includes
#include "msxml6.tlh"

_COM_SMARTPTR_TYPEDEF(IXMLDOMDocument  , __uuidof(IXMLDOMDocument));   // Defines the smart pointer type IXMLDOMDocumentPtr
_COM_SMARTPTR_TYPEDEF(IPersistStream   , __uuidof(IPersistStream));    // Defines the smart pointer type IPersistStreamPtr.
_COM_SMARTPTR_TYPEDEF(IXMLDOMParseError, __uuidof(IXMLDOMParseError)); // Defines the smart pointer type IXMLDOMParseErrorPtr.


/************************************************
*
*                 WINDOWS ERROR
*
************************************************/

/// Get formatted text for Windows error code (such as HRESULT).
std::string WinErr(unsigned int nErrCode, const std::vector<std::string> & moduleslist)
{
  LPVOID lpMsgBuf = nullptr;

  //unsigned int facility = HRESULT_FACILITY(nErrCode); // This is not used. Mainly here for debug purpose.

  // First try to get message from system (using NULL for lpSource)
  ::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK,
    NULL,
    nErrCode,
    MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
    (LPTSTR) &lpMsgBuf,
    0,
    NULL);

  if( !lpMsgBuf )
  {
    // Iterate through list of modules which might contain the error text that can be used by FormatMessage...
    for( auto it = moduleslist.begin(); !lpMsgBuf && it != moduleslist.end(); ++it )
    {
      HINSTANCE hInst = 0;
      bool freelibrary(false);
      if( !it->empty() )
      {
        // Get module instance (only succeeds if the module already is loaded into process)
        hInst = GetModuleHandle(it->c_str());
        if( !hInst )
        {
          // The module is not already loaded so do it and set freelibrary flag so that we remember to free (unload) library after use.
          hInst = LoadLibrary(it->c_str());
          freelibrary = hInst != 0;
        }
      }

      ::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK | FORMAT_MESSAGE_FROM_HMODULE,
        hInst,
        nErrCode,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
        (LPTSTR) &lpMsgBuf,
        0,
        NULL);

      if( freelibrary )
      {
        bool ok = FreeLibrary(hInst) == TRUE;
        assert(ok);
      }
    }
  }

  std::ostringstream ostr;

  if( lpMsgBuf )
  {
    ostr << static_cast<const char*>(lpMsgBuf);
  }
  else
  {
    // The error is known - but it is the error text which is unknown! FormatMessage() could most probably
    // not load the error text because the message table module is not known by it. This happens most often 
    // if the error code is returned from a COM call and hence the actual descriptive error text is present 
    // in whatever COM server, i.e. exe/dll module, containing the message table for the COM Server. If that 
    // module is unknown to FormatMessage(), the text returned is empty.
    //
    // Example: Some COM calls through to MSXML may return error codes which cannot be deduced by FormatMessage(),
    //          unless of course FormatMessage() get the lpSource (2nd argument) to the HMODULE (e.g. MSXML6.dll).
    //          Now, for MSXML error, we in EBC instead usually calls XMLDomParseErrStr().
    //
    // As developer you may try to find the text manually by using the Error Lookup (ERRLOOKUP.EXE) utility in 
    // Visual Studio's "Tools | Error Lookup" menu item. http://msdn2.microsoft.com/en-us/library/akay62ya(VS.80).aspx
    // Note that it may be necessary that you manually add the missing module, try e.g. to add the 
    // "C:\Windows\System\Msxml6r.dll", which contains the message table for MSXML.
    // In general though it may not always be easy to find the correct message table.

    ostr << "Could not format error text for error code 0x" << std::hex << nErrCode << ".";
  }

  // Free the buffer
  LocalFree(lpMsgBuf);

  return ostr.str();
}

/************************************************
*
*             XML DOM PARSE ERROR
*
************************************************/

// Get formatted text for XMLDOM Parse error for given XMLDOM Document.
std::string XMLDomParseErrStr(IXMLDOMDocumentPtr idoc)
{
  std::ostringstream ostr;

  IXMLDOMParseErrorPtr ixmlerr;
  idoc->get_parseError(&ixmlerr);

  //long errcode = ixmlerr->GeterrorCode();

  long errcode{0};
  HRESULT hr = ixmlerr->get_errorCode(&errcode);
  if( SUCCEEDED(hr) )
  {
    BSTR reason;
    long linepos;
    long lineNumber;
    BSTR srcText;

    ixmlerr->get_reason(&reason);
    ixmlerr->get_line(&lineNumber);
    ixmlerr->get_linepos(&linepos);
    ixmlerr->get_srcText(&srcText);

    ostr << _bstr_t(reason,false) << "\n";
    ostr << "Line " << lineNumber << ", Position " << linepos << "\n\n";
    //ostr << _bstr_t(srcText,false) << "\n";

    //// If( character ) position in line is available...
    //if( linepos>0 )
    //{
    //  ostr << "\n";
    //  // Build marker to the (character) position
    //  for( int i = 0; i < linepos - 1; i++ )
    //    ostr << '-';

    //  ostr << '^';
    //}
  }

  return ostr.str();
}

/************************************************
*
*               STREAM ADAPTER
*
************************************************/

/// Stream adapter allowing to access file as a IStream.
/// This is just to trigger the XML problem via loading through IStream.
class StreamAdapter: public IStream
{
public:
  // IUnknown support
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void __RPC_FAR* __RPC_FAR* ppvObject)
  {
    if( riid == IID_IStream )
    {
      *ppvObject = static_cast<IStream*>(this);
      return S_OK;
    }
    else
    {
      ppvObject = nullptr;
      return E_NOINTERFACE;
    }
  }

  ULONG STDMETHODCALLTYPE AddRef(void) { return 0; }
  ULONG STDMETHODCALLTYPE Release(void) { return 0; };

  // IStream support - for now all return E_NOTIMPL
  HRESULT STDMETHODCALLTYPE Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER __RPC_FAR* plibNewPosition) { std::cout << "Seek\n";  return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE SetSize(ULARGE_INTEGER libNewSize) { std::cout << "SetSize\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE CopyTo(IStream __RPC_FAR* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER __RPC_FAR* pcbRead, ULARGE_INTEGER __RPC_FAR* pcbWritten) { std::cout << "CopyTo\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE Commit(DWORD grfCommitFlags) { std::cout << "Commit\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE Revert(void) { std::cout << "Revert\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) { std::cout << "LockRegion\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) { std::cout << "UnlockRegion\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE Stat(STATSTG __RPC_FAR* pstatstg, DWORD grfStatFlag) { std::cout << "Stat\n"; return E_NOTIMPL; };
  HRESULT STDMETHODCALLTYPE Clone(IStream __RPC_FAR* __RPC_FAR* ppstm) { std::cout << "Clone\n"; return E_NOTIMPL; };

  StreamAdapter(const std::string& filename)
  {
    file.open(filename.c_str(), std::ios::binary);
  }

  ~StreamAdapter()
  {
    file.close();
  }

  /// ISequentialStream::Read() implementation - reading from a stream in an archive.
  HRESULT STDMETHODCALLTYPE Read(void __RPC_FAR* pv, ULONG cb, ULONG __RPC_FAR* pcbRead) override
  {
    if( !file )
      return S_FALSE;

    file.read((char*) pv, cb);
    const auto nread = file.gcount();
    if( pcbRead )
      *pcbRead = nread;

    return nread ? S_OK : S_FALSE;
  }

  /// ISequentialStream::Write() implementation - writing to a stream in an archive.
  HRESULT STDMETHODCALLTYPE Write(const void __RPC_FAR* pv, ULONG cb, ULONG __RPC_FAR* pcbWritten) override
  {
    // Not used
    return S_FALSE;
  }


private:
  std::ifstream file;
};

/************************************************
*
*     COINITIALIZE/COUNITIALIZE RAII HELPER
*
************************************************/

// Dirty quick RAII wrapper for CoInitialize/CoUnitialize 
// Would otherwise recommend to use wil::unique_couninitialize_call
class CoInit
{
public:
  CoInit() { ::CoInitialize(nullptr); }
  ~CoInit() { ::CoUninitialize(); }
};

/************************************************
*
*                     MAIN
*
************************************************/

int main()
{
  CoInit coinit;

  {
    CLSID clsId;

    HRESULT hr = CLSIDFromProgID(L"Msxml2.DOMDocument.6.0", &clsId);
    if( !SUCCEEDED(hr) )
    {
      std::cout << "Error: Cannot create class id for MSXML: " << WinErr(hr, {}) << "\n";
      return 1;
    }

    IXMLDOMDocumentPtr idoc;
    hr = idoc.CreateInstance(clsId, nullptr, CLSCTX_INPROC_SERVER);
    if( !SUCCEEDED(hr) )
    {
      std::cout << "Error: Cannot create XMLDOMDocument instance: " << WinErr(hr, {}) << "\n";
      return 1;
    }

    IPersistStreamPtr ipsi;
    hr = idoc.QueryInterface(IID_PPV_ARGS(&ipsi));
    if( !SUCCEEDED(hr) )
    {
      std::cout << "Error: IPersistStream interface unexpectedly not supported: " << WinErr(hr, {}) << "\n";
      return 1;
    }

    //StreamAdapter stream("sample-good.xml"); // Will succeed without problem.
    StreamAdapter stream("sample-bad.xml");    // Will cause exception/failure to load.

    // Load..
    hr = ipsi->Load(&stream); /// Will throw unexpected exception and error out. Exception thrown at 0x75F3ED42 (KernelBase.dll) in xmlcrash.exe: WinRT originate error - 0x80004005 : 'Unspecified error'.
    if( !SUCCEEDED(hr) )
    {
      std::cout << "Error loading xml document from file using stream adapter.\n";
      std::cout << WinErr(hr, {"msxml6r.dll"}) << "\n";
      std::cout << "DOM Parse Error " << XMLDomParseErrStr(idoc);
      return 1;
    }

  }

  return 0;
}