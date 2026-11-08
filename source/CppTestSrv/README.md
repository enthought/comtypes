# Native COM Test Server (`ComtypesCppTestSrvLib`)

## Purpose

`ComtypesCppTestSrvLib` is a first-party, out-of-process (Local Server) native C++ COM server used as a test double for the `comtypes` test suite.

Its primary purposes are:

- **Validating Behavior Across the Language Boundary**:  
  `comtypes` bridges Python and native COM components. Pure Python unit tests or mocks can only exercise Python-side code and cannot verify whether data types, memory layouts, or pointers are correctly interpreted by native COM code. `ComtypesCppTestSrvLib` provides real native endpoints to verify type conversion, data packing/unpacking, and boundary semantics.

- **Exercising Real COM Marshaling and Runtime Semantics**:  
  Because `ComtypesCppTestSrvLib` runs as an out-of-process server (`server.exe`), communication crosses process boundaries and goes through the Windows COM runtime (OLE Automation marshaler, proxy/stub, and RPC channel). This ensures that tests exercise genuine COM marshaling—including SAFEARRAY layout, `IRecordInfo` handling, and `IDispatch::Invoke` argument passing—rather than bypassing marshaling in-process.

- **Dependency-Free, Deterministic Testing**:  
  Rather than relying on external third-party software (such as Microsoft Office or commercial COM servers) which may not be present on every developer machine or CI runner, `ComtypesCppTestSrvLib` acts as a self-contained, lightweight test double that can be compiled, registered, tested, and unregistered on demand.

---

## Modification Policies

When adding or modifying test doubles in `source/CppTestSrv`, follow these policies:

### 1. Scope and Justification
- **Use only when necessary**:  
  Do not introduce or modify native test doubles if the behavior can be adequately and reliably tested using pure Python. Use `ComtypesCppTestSrvLib` specifically when verifying COM marshaling, type library handling, or native memory/type boundary interactions.
- **Keep test doubles minimal and focused**:  
  Each test double component or interface should address a specific testing objective (e.g., testing SAFEARRAY conversions or record dispatch). Avoid adding unrelated methods or behaviors to an existing class merely to support a new test.

### 2. Interface Stability and Isolation
- **Preserve existing interfaces and ABI**:  
  Do not alter existing IDL interfaces, methods, structs, or coclasses unless fixing an existing defect in that specific test double. Existing regression tests rely on their established contracts and GUIDs.
- **Prefer adding new interfaces or coclasses**:  
  When introducing support for a new COM type or scenario, add a new interface or coclass rather than mutating existing ones. Generate new, unique GUIDs for all new IDL definitions.
- **Design for standard COM semantics**:  
  Follow COM requirements strictly, including proper interface inheritance (`IUnknown`, `IDispatch`), correct IDL attributes (`[dual]`, `[oleautomation]`, etc.), and proper reference counting.

### 3. Memory Management and Robustness
- **Strictly adhere to COM memory conventions**:  
  Callee/caller allocation responsibilities (e.g., `CoTaskMemAlloc`/`CoTaskMemFree`, `SafeArrayCreate`/`SafeArrayDestroy`) must be followed without leaking memory or invalidating pointers.
- **Fail cleanly**:  
  Native methods must validate input parameters, dimensions, and null pointers, returning appropriate `HRESULT` error codes (such as `E_POINTER`, `E_INVALIDARG`, `E_OUTOFMEMORY`) rather than crashing or causing access violations that would terminate the server process unexpectedly.

### 4. Deterministic and Predictable Expectations
- **Keep native logic simple and predictable:**
  Native COM methods do not need complex business logic; echoing or mirroring input values (e.g., accepting a value of a certain type and returning an equivalent value of the same type) is entirely sufficient to test that argument and return-value marshaling work correctly. Meanwhile, Python-side test code (assertions) should use hardcoded literal values to ensure better test maintainability.
- **Make regressions readily detectable:**
  Test doubles should provide predictable, distinct values (e.g., non-zero lower bounds, specific array dimensions, signed/unsigned boundary values) to ensure that partial deserialization, wrong element sizing, or off-by-one errors are caught immediately.

### 5. Coordinated Changes Across the Codebase
- **Synchronize all components**:  
  Any change to `ComtypesCppTestSrvLib` must update all relevant layers consistently:
  - IDL definition (`SERVER.IDL`)
  - C++ implementation and factory registration (`SERVER.CPP`, `MAKEFILE`)
  - Associated Python tests in `comtypes/test/`
- **Maintain multi-architecture compatibility**:  
  The C++ code and build configuration must compile cleanly under MSVC on both 32-bit (`x86`) and 64-bit (`x64`) architectures supported by the CI matrix.
- **Adhere to test skip conventions**:  
  Tests utilizing `ComtypesCppTestSrvLib` must handle the absence of the compiled or registered server gracefully.  
  When `GetModule` or imports fail, tests must catch `(ImportError, OSError)` and skip cleanly, without hiding assertion failures when the server is present.
- **Ensure clean registration and unregistration**:  
  The server must support idempotent registration (`/RegServer`) and unregistration (`/UnregServer`), leaving no orphaned registry entries upon unregistration.
- **Verify in CI**:  
  Because contributors may not always have a complete local MSVC build environment, verify that CI builds and registers the server cleanly and that all tests pass across all matrix combinations.
