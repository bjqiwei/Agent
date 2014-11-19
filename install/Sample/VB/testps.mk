
testps.dll: dlldata.obj test_p.obj test_i.obj
	link /dll /out:testps.dll /def:testps.def /entry:DllMain dlldata.obj test_p.obj test_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del testps.dll
	@del testps.lib
	@del testps.exp
	@del dlldata.obj
	@del test_p.obj
	@del test_i.obj
