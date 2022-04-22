# xmlcrash

Illustrate MSXML6 failure to load XML (IXMLDOMDocument) via IPersistStream.

Contains two sample xml files [sample-bad.xml](sample-bad.xml) and [sample-good.xml](sample-good.xml). The [sample-bad.xml](sample-bad.xml) is having a very long XML attribute value (in this example with some RTF formatted text) which likely is causing the MSXML6 parser problem.

Can be reproduced with MSXML6.DLL version 6.30.22000.593 (part of Windows 11).

## Build

Build using Visual C++ compiler (cl.exe):

    cl /EHsc /Zi /D DEBUG xmlcrash.cpp
  
## Sample Output

    Error loading xml document from file using stream adapter.
    Unspecified error
    DOM Parse Error Unspecified error

    Line 7, Position 264165
    
## Main Code

The main part of the problem are illustrated by the following C++ source code from the [xmlcrash.cpp](xmlcrash.cpp):

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
    hr = ipsi->Load(&stream); // Will throw unexpected exception and error out.
                              // Exception thrown at 0x75F3ED42 (KernelBase.dll) in xmlcrash.exe:
                              // WinRT originate error - 0x80004005 : 'Unspecified error'.
    if( !SUCCEEDED(hr) )
    {
      std::cout << "Error loading xml document from file using stream adapter.\n";
      std::cout << WinErr(hr, {"msxml6r.dll"}) << "\n";
      std::cout << "DOM Parse Error " << XMLDomParseErrStr(idoc);
      return 1;
    }

The `class StreamAdapter` is just a convinient wrapper to allow the sample file to be loaded using `ISequentialStream::Read()`.
