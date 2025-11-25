# LibreOffice XScriptProvider Empty Interface Error - Help Needed

## Problem Summary

We are unable to programmatically execute embedded Python-UNO macros via `XScriptProvider` in an automated testing environment (headless + Xvfb). All attempts to execute scripts fail with:

```
invalid attempt to assign an empty interface of type com.sun.star.script.provider.XScriptProvider!
at ./include/com/sun/star/uno/Reference.hxx:105
at ./scripting/source/provider/MasterScriptProvider.cxx:315
```

**Important**: The same Python macros execute successfully when the ODS file is opened manually in LibreOffice GUI on the same system.

## Environment

- **OS**: Linux (Ubuntu/Debian-based) - `Linux 6.8.0-86-generic`
- **LibreOffice Version**: System-installed (via apt)
- **Python**: 3.11+ with `python3-uno` from system packages
- **UNO Connection**: Socket-based (host=127.0.0.1, port=2002)
- **Display Setup**: Xvfb virtual display (tried displays :99-:200)

## What We're Trying to Do

We have a Python application (`xlsliberator`) that:

1. Converts Excel files (.xlsm) to LibreOffice Calc (.ods)
2. Translates VBA macros to Python-UNO and embeds them in the ODS file
3. Attempts to validate that embedded Python macros can execute programmatically

The embedded Python macros are:
- Syntactically valid Python code
- Properly embedded in `Scripts/python/*.py` within the ODS ZIP structure
- Registered in `META-INF/manifest.xml`
- Referenced in `content.xml` event handlers (e.g., button onClick events)

## The Exact Error

When attempting to execute any embedded Python script via XScriptProvider, we get:

```
com.sun.star.uno.RuntimeException: invalid attempt to assign an empty interface
of type com.sun.star.script.provider.XScriptProvider!
at ./include/com/sun/star/uno/Reference.hxx:105
at ./scripting/source/provider/MasterScriptProvider.cxx:315
```

### Our Code Pattern

```python
import uno
from com.sun.star.script.provider import XScriptProvider, XScript
from com.sun.star.beans import PropertyValue

# Setup (both patterns tried)
# Pattern 1: Headless
libreoffice_process = subprocess.Popen([
    "soffice", "--headless",
    "--accept=socket,host=127.0.0.1,port=2002;urp;",
    "--norestore", "--nofirststartwizard"
])

# Pattern 2: GUI mode with Xvfb (current approach)
xvfb_process = subprocess.Popen([
    "Xvfb", ":110", "-screen", "0", "1280x1024x24", "-nolisten", "tcp"
])
env = os.environ.copy()
env["DISPLAY"] = ":110"
libreoffice_process = subprocess.Popen([
    "soffice",  # No --headless flag
    "--accept=socket,host=127.0.0.1,port=2002;urp;",
    "--norestore", "--nofirststartwizard"
], env=env)

# Connect via UNO
local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local_context
)
uno_url = "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"
remote_context = resolver.resolve(uno_url)
desktop = remote_context.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", remote_context
)

# Open document with macros enabled
load_props = (PropertyValue(Name="MacroExecutionMode", Value=4),)  # ALWAYS_EXECUTE_NO_WARN
doc = desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)

# Attempt to get script provider - THIS IS WHERE IT FAILS
script_provider = doc.getScriptProvider()  # Returns empty interface!

# Alternative attempts (all fail the same way):
# script_provider = remote_context.getValueByName(
#     "/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"
# ).createScriptProvider("")

# Try to execute script
script_uri = "vnd.sun.star.script:MyModule.py$my_function?language=Python&location=document"
script = script_provider.getScript(script_uri)  # Fails here
script.invoke((), (), ())
```

### Observations

1. **`doc.getScriptProvider()` returns an empty interface**: The method succeeds but returns an unusable XScriptProvider reference
2. **Alternative factory approach fails identically**: Using `theMasterScriptProviderFactory` produces the same error
3. **Works in manual GUI**: Opening the same ODS file in LibreOffice GUI (same system) allows Python macros to execute perfectly
4. **Headless vs GUI mode**: Tried both `--headless` and GUI mode with Xvfb - same failure
5. **Macro security**: Set to Low (0) via ConfigurationUpdateAccess

## What We've Tried

### 1. Different LibreOffice Startup Modes
- ✗ Headless mode (`--headless`)
- ✗ GUI mode with Xvfb virtual display
- ✗ Different display numbers (:99-:200)
- ✗ Different screen resolutions (1024x768x24, 1280x1024x24)

### 2. Different Script Provider Access Methods
```python
# Method 1: From document
script_provider = doc.getScriptProvider()  # Empty interface

# Method 2: From remote context
msp_factory = remote_context.getValueByName(
    "/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"
)
script_provider = msp_factory.createScriptProvider("")  # Empty interface

# Method 3: With document as argument
script_provider = msp_factory.createScriptProvider(doc)  # Empty interface
```

### 3. Macro Security Settings
```python
config_provider = remote_context.ServiceManager.createInstanceWithContext(
    "com.sun.star.configuration.ConfigurationProvider", remote_context
)
config_path = PropertyValue()
config_path.Name = "nodepath"
config_path.Value = "/org.openoffice.Office.Common/Security/Scripting"
config_access = config_provider.createInstanceWithArguments(
    "com.sun.star.configuration.ConfigurationUpdateAccess", (config_path,)
)
config_access.setPropertyValue("MacroSecurityLevel", 0)  # Low
config_access.commitChanges()
```

### 4. Document Loading Options
```python
# Tried various combinations:
load_props = (
    PropertyValue(Name="MacroExecutionMode", Value=4),  # ALWAYS_EXECUTE_NO_WARN
    PropertyValue(Name="Hidden", Value=False),
    PropertyValue(Name="ReadOnly", Value=False),
)
```

### 5. Environment Variables
```python
env["DISPLAY"] = ":110"
env["DBUS_SESSION_BUS_ADDRESS"] = ""  # Tried with/without
```

## File Structure of Embedded Macros

The ODS file is a valid ZIP containing:

```
myfile.ods/
├── META-INF/
│   └── manifest.xml          # Contains script entries
├── Scripts/
│   └── python/
│       ├── MyModule.py       # Embedded Python-UNO code
│       └── AnotherModule.py
├── content.xml               # Contains event handlers
├── styles.xml
├── meta.xml
└── settings.xml
```

### manifest.xml entry:
```xml
<manifest:file-entry
    manifest:media-type="application/vnd.sun.star.script"
    manifest:full-path="Scripts/python/MyModule.py"/>
```

### content.xml event handler:
```xml
<form:button form:name="StartButton">
    <script:event
        script:event-name="form:performaction"
        script:language="Python"
        script:macro-name="vnd.sun.star.script:MyModule.py$start_game?language=Python&amp;location=document"
        xlink:href="vnd.sun.star.script:MyModule.py$start_game?language=Python&amp;location=document"
        xlink:type="simple"/>
</form:button>
```

### Example embedded Python macro:
```python
# Scripts/python/MyModule.py
import uno
import unohelper
from com.sun.star.awt import XActionListener

def start_game(event=None):
    """Function called by button click."""
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    cell = sheet.getCellByPosition(0, 0)
    cell.setString("Game Started!")
    return None
```

## Questions for LibreOffice Expert

1. **Is XScriptProvider initialization complete during socket-based UNO connection?**
   - Does it require additional services to be started?
   - Are there initialization steps we're missing?

2. **Is there a difference in script provider setup between interactive GUI and programmatic/headless access?**
   - Does the scripting framework require a "real" X display?
   - Are there specific services that only initialize in GUI mode?

3. **What conditions must be met for `doc.getScriptProvider()` to return a valid interface?**
   - Document state requirements?
   - Component manager registration?
   - Service dependencies?

4. **Is there a way to diagnose why the XScriptProvider interface is empty?**
   - UNO introspection tools?
   - Debug logging we can enable?
   - Service manager queries to check provider availability?

5. **Are Python-UNO macros supposed to work in automated/headless environments?**
   - Known limitations?
   - Alternative approaches for programmatic macro execution?

6. **Could this be related to Python UNO bindings vs LibreOffice version mismatch?**
   - We use system `python3-uno` package
   - Is there a version compatibility check we should perform?

## Reproduction Steps

1. Install LibreOffice and python3-uno:
   ```bash
   apt-get install libreoffice python3-uno xvfb
   ```

2. Create test ODS with embedded Python macro:
   ```bash
   git clone https://github.com/johannhartmann/xlsliberator.git
   cd xlsliberator
   python -m pip install -e .

   # Convert Excel with VBA to ODS with Python macros
   xlsliberator convert tests/data/Tetris.xlsm tests/output/Tetris.ods
   ```

3. Attempt to execute macro programmatically:
   ```python
   from xlsliberator.python_macro_manager import test_all_macros_safe
   test_all_macros_safe("tests/output/Tetris.ods")
   # Fails with XScriptProvider empty interface error
   ```

4. Verify macro works in GUI:
   ```bash
   libreoffice tests/output/Tetris.ods
   # Click "Start" button - Python macro executes successfully
   ```

## Workaround We Currently Use

We skip runtime macro validation and rely on:
1. Syntax validation (AST parsing)
2. Manual testing in GUI
3. Documentation warning users that automated execution testing is unavailable

However, we would prefer to have automated validation for CI/CD pipelines.

## Additional Context

- **Source code**: https://github.com/johannhartmann/xlsliberator
- **Relevant files**:
  - `src/xlsliberator/python_macro_manager.py` (lines 180-220)
  - `src/xlsliberator/uno_conn.py` (UNO connection management)
  - `src/xlsliberator/api.py` (line 382 - TODO documenting this issue)

## What We Need

Guidance on:
1. How to properly initialize XScriptProvider in programmatic/automated mode
2. Alternative approaches for validating embedded Python macros work correctly
3. Whether this is a known limitation or a configuration issue
4. Debug steps to identify why the interface is empty

Any help would be greatly appreciated! We're happy to provide additional logs, test files, or run diagnostic commands.
