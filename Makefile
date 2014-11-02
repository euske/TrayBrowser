# Makefile

DEL=del /f
COPY=copy /y
MT=mt -nologo
CC=cl /nologo
RC=rc /nologo
LINK=link /nologo

CFLAGS=/MD /O2 /GA /Zi
LDFLAGS=/DEBUG /MACHINE:x86 /OPT:REF /OPT:ICF
RCFLAGS=
DEFS_COMMON=/D WIN32 /D UNICODE /D _UNICODE
DEFS_CONSOLE=$(DEFS_COMMON) /D CONSOLE /D _CONSOLE
DEFS_WINDOWS=$(DEFS_COMMON) /D WINDOWS /D _WINDOWS
DEFS=$(DEFS_WINDOWS)
#DEFS=$(DEFS_CONSOLE)
LIBS=
INCLUDES=
TARGET=TrayBrowser.exe

all: $(TARGET)

test: $(TARGET)
	.\$(TARGET)

clean:
	-$(DEL) $(TARGET)
	-$(DEL) *.lib *.exp *.obj *.res *.ilk *.pdb *.manifest

$(TARGET): TrayBrowser.res TrayBrowser.obj AXClientSite.obj ini.obj
	$(LINK) $(LDFLAGS) /manifest /out:$@ $** $(LIBS)
	$(MT) -manifest $@.manifest -outputresource:$@;1

TrayBrowser.cpp: Resource.h AXClientSite.h ini.h
TrayBrowser.rc: Resource.h TrayBrowser.ico
AXClientSite.cpp: AXClientSite.h
ini.c: ini.h

.cpp.obj:
	$(CC) $(CFLAGS) /Fo$@ /c $< $(DEFS) $(INCLUDES)
.rc.res:
	$(RC) $(RCFLAGS) $<
