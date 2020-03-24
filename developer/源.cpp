#define _AFX_SECURE_NO_WARNINGS
#pragma warning(disable : 4996)
#define _CRT_SECURE_NO_WARNINGS

#include<stdio.h>
#include<string.h>
#include<stdlib.h>
#include "sapi.h"
#include "sphelper.h"

//#include <stdio.h>
using namespace std;
#pragma comment(lib,"ole32.lib")
#pragma comment(lib,"sapi.lib")
#define SK if(SUCCEEDED(hr))
#define TEXTMAX 20
#define LEN 100000
#define MAN true
#define WOM false
#define CHNTMP "temp.wav"
#define TXT "材料.txt"
#define CMD "格式.txt"
#define MAT "all.wav"
#define MP3 "听力音频.mp3"
//#define min(_a,_b) ((_a)<(_b)?(_a):(_b))

//char names[3002][25];
char blank[2500000];
int txtnum = 0;
char header[44] =
{ 'R','I','F','F',
0,0,0,0,//+36
'W','A','V','E','f','m','t',' ',
0x10,0,0,0,1,0,1,0,
0,0,0,0,0,0,0,0,//samplerate,2*samplerate
2,0,0x10,0,'d','a','t','a',
0,0,0,0 };
unsigned samplerate = 16000;
unsigned totallength = 0;

FILE* fmat, *fout;
char s[LEN + 1];
char text[TEXTMAX + 1][LEN / 4 + 1];
int dialognum[TEXTMAX + 1] ;

int SaveVoice(const char* Text, const char* FileName, const int speaker)
{
    CComPtr<ISpVoice> m_cpVoice;
	WCHAR Wchar[LEN / 4];
	MultiByteToWideChar(CP_ACP, 0, Text, strlen(Text) + 1, Wchar, (LEN / 4) / sizeof(Wchar[0]));
    //初始化COM接口
    if (FAILED(::CoInitialize(NULL)))return -1;
    HRESULT hr = m_cpVoice.CoCreateInstance(CLSID_SpVoice);
    if (SUCCEEDED(hr))
    {
        USES_CONVERSION;
        CComPtr<ISpStream> cpWavStream;
        CComPtr<ISpStreamFormat> cpOldStream;
        CSpStreamFormat OriginalFmt;
        hr = m_cpVoice->GetOutputStream(&cpOldStream);
        if (hr == S_OK) hr = OriginalFmt.AssignFormat(cpOldStream);
        else hr = E_FAIL;
        if (SUCCEEDED(hr)) hr = SPBindToFile(FileName, SPFM_CREATE_ALWAYS, &cpWavStream, &OriginalFmt.FormatId(), OriginalFmt.WaveFormatExPtr());
        if (SUCCEEDED(hr)) hr = m_cpVoice->SetOutput(cpWavStream, TRUE);
        if (SUCCEEDED(hr))
        {
            ISpObjectToken* p = NULL;
            if(speaker==3)         hr = SpFindBestToken(SPCAT_VOICES, L"language=804", L"Name=Microsoft Huihui Desktop", &p);
			else if (speaker == 1) hr = SpFindBestToken(SPCAT_VOICES, L"language=409", L"Name=Microsoft David Desktop", &p);
			else if (speaker == 2) hr = SpFindBestToken(SPCAT_VOICES, L"language=409", L"Name=Microsoft Zira Desktop", &p);
            if (SUCCEEDED(hr))m_cpVoice->SetVoice(p);
            if(speaker==1||speaker==2) m_cpVoice->SetRate(-1);
			else  m_cpVoice->SetRate(-1);
            m_cpVoice->SetVolume(100);
            hr = m_cpVoice->Speak(Wchar, SPF_ASYNC | SPF_IS_XML, 0);
        }
        m_cpVoice->WaitUntilDone(INFINITE);
        cpWavStream.Release();
        m_cpVoice->SetOutput(cpOldStream, FALSE);
        m_cpVoice.Release();
        m_cpVoice = NULL;
    }
    ::CoUninitialize();
    return 0;
}
void syneng(const char* sp, bool sex, int txt, int cts)
{
	char filename[25],mater[110000];
	sprintf(filename, "T%dS%d.wav", txt, cts);
	strcpy(mater, "<pitch absmiddle=\"8\"/>");
	strcat(mater, sp);
	dialognum[txt] = max(dialognum[txt], cts);
	if (sex == MAN)
	{
		SaveVoice(mater, filename, 1);
	}
	else
	{
		SaveVoice(mater, filename, 2);
	}
}
void getspeech(int textnum)
{
	const char* const M = "M: ", * const W = "W: ";
	//char speech[LEN / 4 + 1];
	char* now = text[textnum], * man = NULL, * woman = NULL;
	int cts = 1;
	txtnum = max(txtnum, textnum);
	printf("合成第 %d 段材料中……\n", textnum);
	for (;;)
	{
		man = strstr(now, M);
		woman = strstr(now, W);
		if (man == woman)
		{
			syneng(now, WOM, textnum, cts++); return;
		}
		else if (man == NULL)
		{
			syneng(woman + strlen(W), WOM, textnum, cts++); return;
		}
		else if (woman == NULL)
		{
			syneng(man + strlen(M), MAN, textnum, cts++); return;
		}
		else if (woman < man)
		{
			*man = 0;
			syneng(woman + strlen(W), WOM, textnum, cts++);
			*man = 'M';
			now = man;
			continue;
		}
		else if (man < woman)
		{
			*woman = 0;
			syneng(man + strlen(M), MAN, textnum, cts++);
			*woman = 'W';
			now = woman;
			continue;
		}
	}
}
void splittext()
{
	char* now = s, * next = NULL, * cp = NULL, * l = NULL;
	char temp[12];
	int textnum = 1;
	for (;; textnum++)
	{
		if (now == NULL) 
			return;
		sprintf(temp, "Text %d", textnum);
		cp = strstr(now, temp);
		if (cp == NULL)
			return; 
		else cp = (now += strlen(temp) + 1);
		sprintf(temp, "Text %d", textnum + 1);
		next = strstr(now, temp);
		//copy
		for (l = text[textnum]; (next != NULL && cp < next) || (next == NULL && *cp != 0); l++, cp++)*l = *cp;
		*l = 0;
		//printf("%s\n",text[textnum]);
		getspeech(textnum);
		now = next;
	}
}
void fillheader()
{
	fwrite(header, 44, 1, fout);
}
void makesilence(unsigned seconds)
{
	unsigned len;
	len = seconds * samplerate * 2;
	totallength += len;
	fwrite(blank, len, 1, fout);
}
void addwav(const char* name)
{
	char x[10000];
	unsigned len;
	FILE* fwav;
	fwav = fopen(name, "rb");
	if (fwav == NULL)return;
	fseek(fwav, 42, SEEK_SET);
	fread(&len, 4, 1, fwav);
	totallength += len;
	while (len >= 10000) 
	{
		fread(x, 10000, 1, fwav),
		fwrite(x, 10000, 1, fout);
		len -= 10000;
	}
	fread(x, len, 1, fwav),
	fwrite(x, len, 1, fout);
	fclose(fwav);
}
void synchn(const char* sp)
{
	char mater[110000];
	strcpy(mater, "<pitch absmiddle=\"2\"/>");
	strcat(mater, sp);
	SaveVoice(mater, CHNTMP, 3);
	addwav(CHNTMP);
	remove(CHNTMP);
}
void correct()
{
	*((int*)(header + 24)) = samplerate;
	*((int*)(header + 28)) = 2 * samplerate;
	*((int*)(header + 4)) = totallength + 36;
	*((int*)(header + 40)) = totallength;
	rewind(fout);
	fillheader();
}
int makewav(const char* name, const char* rules)
{
	int i;
	int t;
	char command[1500], dlgname[25];
	//unsigned totallength,len;
	FILE* frule;
	totallength = 0;
	fout = fopen(name, "wb");
	frule = fopen(rules, "r");
	//if (fout == NULL || frule == NULL)return 0;
	printf("正在生成……\n");
	fillheader();
	while (fscanf(frule, "%s", command) != EOF)
	{
		switch (*command)
		{
		case ':':
			synchn(command+1);
			break;
		case 'W':
			t = atoi(command + 1);
			makesilence(t);
			break;
		case 'R':
			t = atoi(command + 1);
			for (i = 1; i <= dialognum[t]; i++)
			{
				sprintf(dlgname, "T%dS%d.wav", t, i);
				addwav(dlgname);
			}
			break;
		default:
			break;
		}
	}
	printf("total:%u\n", totallength);
	correct();
	fclose(fout);
	fclose(frule);
	return 1;
}
void deleteall()
{
	char filename[25];
	int i, j;
	printf("删除文件中……\n");
	for (i = 1; i <= txtnum; i++)
	{
		for (j = 1; j <= dialognum[i]; j++)
		{
			sprintf(filename, "T%dS%d.wav", i, j);
			remove(filename);
		}
	}
}
void makemp3()
{
	char command[100];
	printf("转化为mp3……\n");
	sprintf(command, "lame %s %s", MAT, MP3);
	system(command);
	remove(MAT);
	printf("完成！\n");
}
int _tmain(int argc, _TCHAR* argv[])
{
	size_t l = 0;
	fmat = fopen(TXT, "r");
	memset(blank, 0, sizeof(blank));
	memset(s, 0, sizeof(s));
	memset(text, 0, sizeof(text));
	memset(dialognum, 0, sizeof(dialognum));
	l = fread(s, sizeof(char), LEN, fmat);
	if (l == 0) fputs("Error reading file", stderr);
	else s[l] = 0;
	//
	splittext();
	//getchar();
	makewav(MAT, CMD);
	deleteall();
	makemp3();
}



