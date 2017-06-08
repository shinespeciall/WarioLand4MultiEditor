//Write in VC++6.0

#include <stdio.h>
#include <stdlib.h>         //exit() is in the header
#include <string.h>
#include <string>

using namespace std;

#define u32 unsigned int
#define u8 unsigned char

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
	//Get all the pointers point to Level Pointers and Flags Stream
	fseek(gbafp,7926400L,0);     //78F280
	u8 ptrStream[96];
	unsigned int i,j;
	for(i=0;i<96;i++)
		fread(&ptrStream[i],sizeof(u8),1,gbafp);
	u32 pointers[24];
	for(i=0;i<96;i+=4)
		pointers[i/4]=(ptrStream[i+3]<<24)+(ptrStream[i+2]<<16)+(ptrStream[i+1]<<8)+ptrStream[i]-134217728L;
	//Then Load Room number after 639068 for each level
	u8 Bytes[288];
	rewind(gbafp);
	fseek(gbafp,6525032L,0);     //639068
	for(i=0;i<288;i++)
		fread(&Bytes[i],sizeof(u8),1,gbafp);

	unsigned char B[24][704];        //For Save all the Bytes
	unsigned int LevelIndex[17]={17,17,17,17,17,17,17,17,17,17,17,17,17,17,17,17};
	unsigned int n=0;
	for(i=0;i<24;i++)
	{
		rewind(gbafp);
		fseek(gbafp,pointers[i],0);
		fread(&B[i],704*sizeof(u8),1,gbafp);
	}
	printf("This patch is recommanded to use with WL4 MultiEditor.exe\n");
	printf("Open file successfully.\n");
	printf("The following level probably can be maximized\n");
	for(i=0;i<17;i++)                //for The 17 Level and the No.24 can be check somewhere else
	{
		for(j=0;j<17;++j)
			if( ((unsigned int) Bytes[12*j]==i) & ((unsigned int) Bytes[12*j+1]<16) )
				break;
		if(j!=17)             //probably there is no way to judge if the Data areas of Levels have been maximized, is there?
		{
			printf("%4d",(unsigned int) Bytes[12*j]);     //temporary output for check
			LevelIndex[n]=j;
			n++;
		}
	}
	printf("\nInput one number stand for the level whose data area you want to maximize (end with Enter): ");
	scanf("%d", &j);
	a=getchar();
	for(i=0;i<17;++i)
		if(LevelIndex[i]==j)
			break;
	if(i==16)
	{
		printf("Wrong Index or don't need to maximize tha level!\tPress Enter to exit...");
		a=getchar();
		exit(0);
	}
	printf("Level%3d, start processing...",(unsigned int) LevelIndex[i]);     //temporary output for check

	fclose(gbafp);
	printf("\nPress Enter to exit...");
	a=getchar();
	return 0;
}
