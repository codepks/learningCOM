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
