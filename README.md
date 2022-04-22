# xmlcrash

Illustrate MSXML6 failure to load XML (IXMLDOMDocument) via IPersistStream.

Contains two sample xml files `sample-bad.xml` and `sample-good.xml`. The `sample-bad.xml` is having a very long XML attribute value which likely is causing the MSXML6 parser problem.

Can be reproduced with MSXML6.DLL version 6.30.22000.593 (part of Windows 11).

## Build

Build using Visual C++ compiler (cl.exe):

    cl /EHsc /Zi /D DEBUG xmlcrash.cpp
  
## Sample Output

    Error loading xml document from file using stream adapter.
    Unspecified error
    DOM Parse Error Unspecified error

    Line 7, Position 264165
