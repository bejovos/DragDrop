# Drag & Drop

## Setup:
- Add library path to system path.
- Add admin privileges to DragDrop32.exe / DragDrop64.exe.
- Add file association.

## Usage modes:
  
### Drag & Drop into everything else
```
> DragDrop32.exe "path_to_application" "%1"
```

### Drag & Drop into everything with cmd args
```
> DragDrop32.exe "path_to_application" "%1" "arguments"
```

### Drap & Drop into Visual Studio
```
> DragDrop32.exe "path_to_devenv_exe" "%1" VS
```

This mode requires **DebugAddin** to be installed.

## Known issues:  
 
Requires write access to folder with target application.