#include <windows.h>

/* Error messages */
char err0[] = "-ERR: Visual Basic 6 not installed";
char err1[] = "-ERR: VB6.EXE path not found";
char err2[] = "-ERR: LNK.EXE not found";
char err3[] = "-ERR: Not enough heap memory";
char err4[] = "-ERR: Error spawning LNK.EXE";

char buf[1024];
WIN32_FIND_DATA fdata;
int copy_LNK;

int ep(void){
STARTUPINFO sinfo;
PROCESS_INFORMATION pinfo;
HKEY regKey;
HANDLE hFind, hOut = GetStdHandle(STD_OUTPUT_HANDLE);
void* pMem;
char* cur;
char* start;
unsigned int stage, isuFMOD, isVB6;

	hFind = FindFirstFile("LNK.EXE",&fdata);
	if(hFind != INVALID_HANDLE_VALUE) FindClose(hFind);
	else{
		/* LNK.EXE not found
		   Maybe MSDEV has spawned linker in the sources directory
		   In such a case we need to bring LNK.EXE to the same directory */
		if(RegOpenKeyEx(HKEY_CLASSES_ROOT,
			"VisualBasic.Project\\shell\\open\\command",
			0,KEY_QUERY_VALUE,&regKey) != ERROR_SUCCESS){
			/* VB6 not installed */
			WriteFile(hOut,err0,sizeof(err0) - 1,&stage,0);
			return -1;
		}
		stage = sizeof(buf) - 1;
		isVB6 = RegQueryValueEx(regKey,"",0,0,buf,&stage);
		RegCloseKey(regKey);
		if(isVB6 != ERROR_SUCCESS){
			/* VB6 path not found */
			WriteFile(hOut,err1,sizeof(err1) - 1,&stage,0);
			return -1;
		}
		start = buf + stage;
		while(start > buf && *start != '\\') start--;
		cur = buf;
		if(*cur == '"') cur++;
		*(long*)start = 'KNL\\';
		*(long*)(start + 4) = 'EXE.';
		*(start + 8) = 0;
		if(!CopyFile(cur,"LNK.EXE",0)){
			/* LNK.EXE not found */
			WriteFile(hOut,err2,sizeof(err2) - 1,&stage,0);
			return -1;
		}
		copy_LNK++;
	}
	start = cur = GetCommandLine();
	pMem = HeapAlloc(GetProcessHeap(),0,strlen(cur) + 24);
	if(!pMem){
		/* Not enough heap memory */
		WriteFile(hOut,err3,sizeof(err3) - 1,&stage,0);
		return -1;
	}
	*((unsigned int*)pMem) = 'KNL';
	stage = 0;
	isuFMOD = 0;
	isVB6 = 0;
	/* Parse the command line */
	while(*cur){
		if(stage){
			if(!isuFMOD &&
				(*((unsigned int*)cur) == 'v_Fu' ||
				*((unsigned int*)cur) == 'v_fu' ||
				*((unsigned int*)cur) == 'V_FU')){
				/* Lookup uFMOD */
				cur += 4;
				if(*cur != 'b' && *cur != 'B') continue;
				cur++;
				if(*((unsigned int*)cur) == 'JBO.' ||
					*((unsigned int*)cur) == 'jbo.'){
					/* GOTCHA! */
					start = (char*)(cur - start);
					cur += 3;
					isuFMOD++;
				}
			}else if(*((unsigned int*)cur) == 'v__:' &&
					*((unsigned int*)cur + 1) == ' Sab'){
					cur += 7;
					isVB6++;
			}
		}else if(*((unsigned int*)cur) == 'KNIL' ||
				*((unsigned int*)cur) == 'knil' ||
				*((unsigned int*)cur) == 'kniL'){
				/* Skip the exe name */
				cur += 4;
				if(*cur == '.') cur += 4;
				if(*cur == '"') cur++;
				if(*cur != ' ') continue;
				stage++;
				start = cur;
				strcpy((char*)pMem + 3,start);
		}
		cur++;
	}
	stage = 0;
	/* Modify the command line only when invoked from VB6
	   and with a reference to uFMOD */
	if(isuFMOD && isVB6){
		cur = (char*)pMem + 3 + (unsigned int)start;
		*((unsigned int*)cur - 1) = 'domf';
		*((unsigned int*)cur)     = 'bil.';
		strcat(cur," /OPT:NOREF /OPT:nowin98");
	}
	/* Run the real linker */
	GetStartupInfo(&sinfo);
	if(CreateProcess(0,pMem,0,0,1,0,0,0,&sinfo,&pinfo)){
		CloseHandle(pinfo.hThread);
		WaitForSingleObject(pinfo.hProcess,INFINITE);
		GetExitCodeProcess(pinfo.hProcess,&stage);
		CloseHandle(pinfo.hProcess);
	}else{
		/* Couldn't spawn the linker */
		WriteFile(hOut,err4,sizeof(err4) - 1,&stage,0);
		return -1;
	}
	HeapFree(GetProcessHeap(),0,pMem);
	return stage;
}

/* Program entry point */
void start(void){
int c, i = ep();
	if(copy_LNK){
		/* Try to delete LNK.EXE */
		for(c = 0; c < 8; c++){
			Sleep(128);
			if(DeleteFile("LNK.EXE")) break;
		}
	}
	ExitProcess(i);
}
