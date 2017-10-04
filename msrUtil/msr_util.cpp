#include <windows.h>
#include <string>
#include "../dictationbridge-core/client/client.h"

bool isValid(char c){
	if(' ' == c) return false;
	if('|' == c) return true;
	if('_' == c) return true;
	return isalpha((int)c) || isdigit((int)c);
}

int CALLBACK WinMain(_In_ HINSTANCE hInstance,
	_In_ HINSTANCE hPrevInstance,
	_In_ LPSTR lpCmdLine,
	_In_ int nCmdShow) {
		int i=0;
		while(lpCmdLine[i] != '\0'){
			if(!isValid(lpCmdLine[i++])) return 1;
		}
		DB_SendCommand(lpCmdLine);
		return 0;
}
