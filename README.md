# Basics
## CComPtr vs CComQUIPtr

### CComPtr
	1. CComPtr is a smart pointer that is used for managing the lifetime of a COM interface pointer.
	2. It is used when you need a simple pointer management for a COM object.
	3. It does not support querying for different interfaces; it only holds a single interface pointer.


### CComQIPtr
	1. CComQIPtr is also a smart pointer, but it specifically extends CComPtr with additional functionality to perform "QueryInterface" operations.
	2. It is for cases when you want to call QueryInterface() in a convenient manner to know whether an interface is supported:
IInterface1* from = ...
CComQIPtr<IInterface2> to( from );
if( to != 0) {
   //supported - use}
This way you can request an interface from a pointer to any (unrelated) COM interface and check whether that request succeeded.
CComPtr is used for managing objects that surely support some interface. You use it as a usual smart pointer with reference counting. It is like CComQIPtr, but doesn't allow the usecase described above and this gives you better type safety.

# Source and Sink
## Source and Sink

	1. Sources are the places where events are invoked
	2. Sinks are the places where event callbacks are caught
	3. Coclass is the code behind

## Copilot

**Callback Basics** <br>
	1. In COM, callbacks allow a component (the “source”) to notify another component (the “sink”) about events or changes.
	2. The source defines an interface with callback methods, and the sink implements that interface to receive notifications.
	3. The source maintains a reference to the sink and invokes its callback methods when needed.

**Creating the Source (Coclass)** <br>
	1. In your IDL file (typically MyComponent.idl), define the callback interface (e.g., IMyCallback).
	2. Make sure the callback interface belongs to a coclass (the source). This coclass can be any existing one.
	3. Add a member variable (a pointer) of the callback interface to the coclass (the source). This member will hold the reference to the sink.
	4. Implement a method (e.g., Advise) in the working interface of the coclass. This method sets the member pointer to the sink.

Example IDL for Source:
// MyComponent.idl
interface IMyCallback : IUnknown {
    HRESULT OnEvent();
};

coclass MySource {
    interface IMyCallback* m_pSink; // Pointer to the sink
    HRESULT Advise([in] IMyCallback* pSink);
};

Creating the Sink (Another Coclass):
	1. Create another coclass (the sink) that implements the IMyCallback interface.
	2. Use the “Implement Interface Wizard” (e.g., in Visual Studio) to generate the implementation for the callback methods.
	3. This sink coclass will handle the actual callback logic.

Example C++ Implementation for Sink:
// MySink.cpp
class MySink : public IMyCallback {
public:
    STDMETHODIMP OnEvent() {
        // Handle the event here
        return S_OK;
    }
};

// In your main code:
MySource source;
MySink sink;
source.Advise(&sink); // Set the sink reference

Usage:
	1. When an event occurs in the source (e.g., data update), call the sink’s callback method (e.g., OnEvent).
	2. The sink’s implementation of OnEvent will be executed.

AT_ADVISE is linking SINK and Source File.

# Keywords
Surfscan-ADE - COM (sharepoint.com)

## Variant_t
_variant_t is a C++ class that encapsulates the VARIANT data type used in COM. It manages resource allocation and deallocation and makes function calls to VariantInit and VariantClear as needed

## _bstr_t 
_bstr_t is another C++ class that encapsulates the BSTR data type, which is a string type used in COM. Similar to _variant_t, it manages resource allocation and deallocation through function calls to SysAllocString, SysFreeString, and other BSTR APIs when appropriate

## LPCTSTR 
LPCTSTR is a type definition used in the Windows API to represent a “Long Pointer to a Constant TCHAR STRing”. It’s part of a set of type definitions that allow developers to write code that can work with both ANSI and Unicode strings. <br>

 In Unicode mode, 
	1. LPCTSTR would be equivalent to const wchar_t*
	2. ANSI mode, it would be const char*
	
## ANSI Mode
ANSI stands for American National Standards Institute, which developed a set of standards for character encoding. In ANSI mode, characters are encoded using single-byte character sets (SBCS), where each character is represented by a single byte. This limits the number of possible characters to 256, which is sufficient for representing the characters of Western languages but not for languages with larger character sets, such as Chinese or Japanese.

## UNICODE Mode
Unicode is a universal character encoding standard that provides a unique number for every character, no matter the platform, program, or language. Unicode can represent a much larger number of characters because it uses more bytes per character. <br> It can use fixed-width encoding (like UTF-16, where each character is typically 2 bytes) or variable-width encoding (like UTF-8, where characters can be 1 to 4 bytes). This allows Unicode to support virtually all characters and symbols from all writing systems around the world, including emojis.


#define __T(x) L ## x
L is a prefix that, when placed before a string literal, turns it into a wide-character string literal in C++. For example, L"Hello" is a wide-character string literal equivalent to wchar_t*. <br>

"##" is the token-pasting operator. It concatenates two tokens into a single token without any whitespace between them.

In Unicode mode, this would expand to:
wchar_t* str = L"Hello";

## HANDLE 
It is a commonly used type defined in winnt.h. It represents a generic handle to an object and is often used for synchronization, file I/O, and other operations. When you work with Windows APIs, you’ll encounter HANDLE as a way to interact with different system resources

# CoCreateInstance

1. It is way to create COM object
2. _variant_t and _bstr_t are provided by the compiler as COM support classes and get used when you use constructs like #import . You can use them if you like.
3. CComVariant and CComBSTR are provided by the ATL libraries.

## Note
Before the client uses any features of the COM API, he must initialize the COM services using CoInitialize. Once this is done, the client must apply the function CoGetClassObject to obtain the required interface of the class factory component (the specified identifier CLSID of the component, which will be discussed below). <br> Immediately after receipt of the factory customer classes, it calls CreateInstance() to create an actual instance of the BVAA class and dismisses the class factory interface. 

# COM BKM

1. When you call COM library functions such as CoCreateInstance passing a guid constant such as "CLSID_CoCar" or "IID_IRegistration", the compiler converts the guid contstant to the real guid by looking up the guid definition file *_i.c
2. Checkout how to create a COM component in the file itself
3. Topics to learn
   - .tlb
   - .idl
   - lib.dll
5. IDL File : Instead of defining interfaces in .h file, you define them in IDL file which on MIDL compilation generates .h file and a guid file called i.c file fo C++ clients and   which can be understood by many lanugages <br>

 Suppose the IDL file name is Shapes.idl, the files generated by MIDL are:
	1.	Shapes.h	Interface definition
	2.	Shapes.tlb	binary type library
	3.	Shapes_i.c	definition of all guids
	5. .tlb file : When you #import a type library say Shapes.tlb, the compiler will generate two C++ files: Shapes.tlh and Shapes.tli
	6. My C# .NET project in which few files were COM visible generated .tlb files, which are required for COM visibility

## ATL
	1. Inheriting interfaces happens through BEGIN_COM_MAP
	BEGIN_COM_MAP(CAny)
	    COM_INTERFACE_ENTRY(IMyInterf)  // step 2
	END_COM_MAP()



# Testcase usage
```
CoInitialize(nullptr);
```
This line initializes the COM library for the thread. nullptr is passed to indicate that we are not using any specific application concurrency model. It's essential to call this function before using any COM components.

```
CComPtr <IKTPersistenceAgent> pa = nullptr;  
HRESULT hr = pa.CoCreateInstance(_T("SPX.DFNorm.PA"));
```
1. The code above it for interfacing with C# written object
2. This block creates an instance of the IKTPersistenceAgent COM interface. The CComPtr is a smart pointer that automatically manages the memory of the COM object referenced.
3. CoCreateInstance() attempts to create an instance of the specified COM class (SPX.DFNorm.PA), which is likely a persistence agent used for storage or other persistence-related functions.

```
CComPtr<IComResCalDFNormFactory> roFactory;  
roFactory.CoCreateInstance(CLSID_SPResCalibDFNormFactory);
```
1. This section is similar to the earlier code for creating the persistence agent. It creates an instance of the IComResCalDFNormFactory, which is likely a factory for creating response calibration objects.
2. CLSID_SPResCalibDFNormFactory indicates the specific class ID for the response calibration factory. The validity of the roFactory object is then checked. If it was not created successfully, an error is printed, and the program exits.

# Consuming C# class in C++
1. To consume C# class in C# you need .tldb file generated from the C# class which contains the definitions
2. Mention the .tlb file in the common.h file
Use CComPtr and ATL libraries to 

## importing
```
#import "SPResFlyHeight.tlb" no_namespace ,named_guids,raw_interfaces_only
```
Explaination
1. **#import** :  This preprocessor directive is used to import a type library (TLB) that contains definitions for COM interfaces, enums, and structures. It allows the C++ compiler to generate header files and implement classes that map to COM interfaces defined in the TLB.
2. "SPResFlyHeight.tlb" : This specifies the path to the Type Library file that you want to import. In this case, it's a file named SPResFlyHeight.tlb. Type Libraries usually contain metadata about the COM components, such as the definitions of interfaces, classes, and their methods.
3. **no_namespace:** This option tells the compiler not to place the imported classes, interfaces, and enums into a namespace.
 - By default, #import would create a specific namespace based on the project name or the library name.
 - However, using no_namespace means that the identifiers will be placed directly into the global namespace, allowing you to use them without needing to qualify them with a namespace.
4.** named_guids:** 
- This option generates named GUIDs for the COM interfaces and classes.
- Normally, the GUIDs are accessible via their structured representations like IID_IComResFlyHeight, but named_guids creates symbolic constants that can be used in your code more readably, such as CLSID_SpResFlyHeightFactory instead of needing to reference the GUID directly.
5. **raw_interfaces:** This instructs the compiler to generate raw interface definitions rather than smart pointers or other wrappers.
  - Raw interfaces allow for more direct access to the underlying COM methods but require careful handling of reference counting and memory management.
  - This can be useful for performance-critical sections of code or when working directly with lower-level COM APIs.
