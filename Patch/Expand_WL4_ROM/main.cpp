//Write in VC++6.0
//draw one gba file to the icon of the application to start patching
//iF you want to compile it yourself, please change the value here: project -> (setting for debug or release) ->setting -> link -> category: output -> stack allocations -> reserve
//and put a big number like 0x10000000 to make the program be possible to use a big array 

#include <stdio.h>
#include <stdlib.h>         //exit() is in the header
#include <string>

using namespace std;

int main(int argc,char *argv[])
{
	unsigned char a;
	if(argc>2)
	{
		printf("Wrong arguments or too many files, you can only draw them in one by one!\nPress Enter to Exit...");
		a=getchar();
		exit(0);
	}
	if(argc==1)
	{
		printf("No files loaded, draw one to the icon of the application!\nPress Enter to Exit...");
		a=getchar();
		exit(0);
	}
	string FileType;
	FileType=argv[1];
	
	if( (FileType.find(".gba",0)+4)!=FileType.length() )
	{
		printf("Wrong File Type!\nPress Enter to Exit...");
		a=getchar();
		exit(0);
	}

	FILE *gbafp;
	if((gbafp=fopen(argv[1],"r+b"))==NULL)
	{
		printf("Cannot Open file!\nPress Enter to Exit...");
		a=getchar();
		exit(0);
	}
	//start to append
	printf("The file will be expand to 32 MB.Please wait...");
	unsigned char B[33554432];  //32 MB
	unsigned int i;
	rewind(gbafp);
	for(i=0;i<8388608;i++)
		fread(&B[i],sizeof(char),1,gbafp);
	fclose(gbafp);
	printf("\nRead file, Please wait...");
	for(i=8388608;i<33554432;i++)
		B[i]='\xFF';
	if((gbafp=fopen(argv[1],"wb"))==NULL)
	{
		printf("Cannot Open file!\nPress Enter to Exit...");
		a=getchar();
		exit(0);
	}
	for(i=0;i<33554432;i++)
	fwrite(&B[i],sizeof(char),1,gbafp);
	fclose(gbafp);
	gbafp=NULL;
	printf("\nFinish!\tPress Enter to exit...");
	a=getchar();
	return 0;
}
