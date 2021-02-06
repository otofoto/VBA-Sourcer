//*********************************************************************
// Демонстрация поиска неполиморфных вирей в документах
// Компилятор Borland C/C++ v5.02
// (c) Климентьев К. aka DrMAD, Самара 2003
//*********************************************************************
#include "windows.h"
#include "ole2.h"
#include "iostream.h"

#define MAXBUF 0xFFFF
#define NAMELEN 256

char Buf[MAXBUF];  // Буфер под поток
char SBuf[MAXBUF]; // Буфер под распакованный текст

/* Поиск чунка внутри потока */
ULONG search_chunk( char *Buf, ULONG streamlen)
{
 ULONG ch_pos; // Позиция чунка внутри потока

 if ((Buf[0]==1)&&(Buf[1]==0x16)&&(Buf[2]==1))
  {
   ch_pos = streamlen/2;
   while ((ch_pos<(streamlen-3)) && ((Buf[ch_pos]!=1)||
         ((Buf[ch_pos+2]&0xF0)!=0xB0))) ch_pos++;
   if (ch_pos < (streamlen-3)) return ch_pos;
  }
  return 0;
}

/* Поиск в исходнике .Import/.Export */
LookAtVirus() {
 int i=0;
 while (SBuf[i]) {
  if ((SBuf[i]=='.')   &&
      (SBuf[i+3]=='p') &&
      (SBuf[i+4]=='o') &&
      (SBuf[i+5]=='r') &&
      (SBuf[i+6]=='t')) cout << endl << "*** Maybe virus! ***" << endl;
  i++;
 }
}

typedef UINT (WINAPI* RTLD)(ULONG, PVOID, ULONG, PVOID, ULONG, PULONG);

/* Распаковка исходника */
view_src( char *StreamName, char *Buf, ULONG ch_pos)
{
 HMODULE h;                // Хэндл библиотеки NTDLL.DLL
 RTLD RtlDecompressBuffer; // Указатель на функцию
 ULONG ulen;               // Длина распакованного фрагмента
 int i=0;

 LoadLibrary("ntdll.dll");
 h = GetModuleHandle("ntdll.dll");
 RtlDecompressBuffer = (RTLD) GetProcAddress(h, "RtlDecompressBuffer");
 if (RtlDecompressBuffer)
  {
   RtlDecompressBuffer(0x2, SBuf, MAXBUF, &(Buf[ch_pos+1]), MAXBUF-ch_pos, &ulen);
   cout << " *** " << StreamName << " *** " << endl;
   while (SBuf[i]) cout << SBuf[i++];
   LookAtVirus();
  }
}

/* Рекурсивный обход потоков */
walk(char *s, LPSTORAGE ls)
{
 OLECHAR FileName[NAMELEN];	// Unicode-имя для structured storage
 char StreamName[NAMELEN];	// ASCII-имя потока
 LPENUMSTATSTG lpEnum=NULL;	// Интерфейс перечислителя
 LPSTORAGE pIStorage=NULL;	// Интерфейс структурированного хранилища
 LPSTORAGE pIStorage2=NULL;	// Интерфейс хранилища нижнего уровня
 LPSTREAM pIStream=NULL;	// Интерфейс потока
 STATSTG stat;		// Очередная запись в каталоге
 ULONG uCount;		// Счетчик перечисления
 ULONG streamlen;	// Реальная длина потока
 ULONG ch_pos;		// Позиция чунка внутри потока

 if (!ls) // Первый вызов
  {
   mbstowcs(FileName, s, 256);
   StgOpenStorage(FileName,NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,
                  NULL,0,&pIStorage);
   walk("", pIStorage);
  }
 else // Повторный рекурсивный вызов
 {
  ls->EnumElements(0,NULL,0,&lpEnum);
  if (lpEnum)
  while (lpEnum->Next(1,&stat,&uCount)==S_OK)
   {
    if (stat.type==STGTY_STORAGE) // Это хранилище
     {
      ls->OpenStorage(stat.pwcsName,NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,
                      NULL,0,&pIStorage2);
      walk("", pIStorage2);
     }
    else // Это поток
     {
      ls->OpenStream(stat.pwcsName,NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,
                     0,&pIStream);
      pIStream->Read(Buf, MAXBUF, &streamlen);
      wcstombs(StreamName, stat.pwcsName, 256);
      ch_pos = search_chunk(Buf, streamlen);
      if (ch_pos) view_src(StreamName, Buf, ch_pos);
     }
    };
   ls->Release();
  }
}

int main(int argc, char* argv[]) {
 if (argc>1) walk(argv[1],NULL);
 else cout << "Usage: SVIR docfile" << endl;
}
