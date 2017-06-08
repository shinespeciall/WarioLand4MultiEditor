//Write in VC++6.0
//draw one gba file to the icon of the application to start patching

#include <stdio.h>
#include <stdlib.h>         //exit() is in the header
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
	u8 Bytes[348];
	rewind(gbafp);
	fseek(gbafp,6525032L,0);     //639068
	for(i=0;i<348;i++)
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
		for(j=0;j<29;++j)
			if( ((unsigned int) Bytes[12*j]==i) & ((unsigned int) Bytes[12*j+1]<16) )
				break;
		if(j!=29)             //probably there is no way to judge if the Data areas of Levels have been maximized, is there?
		{
			printf("%4d",(unsigned int) Bytes[12*j]);     //output for checking
			LevelIndex[n]=i;
			n++;
		}
	}
	for(j=0;j<29;j++)                //for the No.24
	{
		if( ((unsigned int) Bytes[12*j]==23) & ((unsigned int) Bytes[12*j+1]<16) )
		{
			printf("%4d",(unsigned int) Bytes[12*j]);
			LevelIndex[n]=23;
			n++;
		}
	}
	printf("\nInput one number stand for the level whose data area you want to maximize (end with Enter): ");
	scanf("%d", &j);
	a=getchar();
	for(i=0;i<17;++i)
		if(LevelIndex[i]==j)
			break;
	if(i==17)
	{
		printf("Wrong Index or don't need to maximize tha level!\tPress Enter to exit...");
		a=getchar();
		fclose(gbafp);
		exit(0);
	}
	printf("Level%3d, start processing...\n",(unsigned int) LevelIndex[i]);     //output for checking
	n=(unsigned int) LevelIndex[i];      //n=level_Index

	for(j=0;j<17;++j)
		if((unsigned int) Bytes[12*j]==n)
			break;
	i=(unsigned int) Bytes[12*j+1];

	unsigned char New_Bytes[704];
	for(j=0;j<(44*i);j++)
		New_Bytes[j]=B[n][j];
	j--;
	for(i=44*j;i<704;i++)
		New_Bytes[i]=(unsigned char) 241;
	//read from 78F970 to find space
	rewind(gbafp);
	fseek(gbafp,7928176L,0);
	u8 SpaceOffset[4];
	u32 SpaceOffset2,offset2;
	for(i=0;i<4;i++)
		fread(&SpaceOffset[i],sizeof(u8),1,gbafp);
	SpaceOffset2=(unsigned int) (SpaceOffset[0]<<24)+(unsigned int) (SpaceOffset[1]<<16)+(unsigned int) (SpaceOffset[2]<<8)+(unsigned int) (SpaceOffset[3]);
	if(SpaceOffset2==4294967295)
		SpaceOffset2=7928192;
	if(SpaceOffset2%4!=0)
		SpaceOffset2=(SpaceOffset2/4)*4+4;
	offset2=SpaceOffset2;
	SpaceOffset2=SpaceOffset2+134217728;         //add 08000000
	//rewrite the pointer in small endian
	rewind(gbafp);
	fseek(gbafp,(7926400+4*n),0);
	fwrite(&SpaceOffset2,sizeof(u32),1,gbafp);
	//rewrite the offset in small endian
	SpaceOffset2=SpaceOffset2+704-134217728;
	SpaceOffset[0]=SpaceOffset2>>24;
	SpaceOffset[1]=(SpaceOffset2<<8)>>24;
	SpaceOffset[2]=(SpaceOffset2<<16)>>24;
	SpaceOffset[3]=(SpaceOffset2<<24)>>24;
	SpaceOffset2=(unsigned int) (SpaceOffset[0])+(unsigned int) (SpaceOffset[1]<<8)+(unsigned int) (SpaceOffset[2]<<16)+(unsigned int) (SpaceOffset[3]<<24);
	rewind(gbafp);
	fseek(gbafp,7928176L,0);   //78F970
	fwrite(&SpaceOffset2,sizeof(u32),1,gbafp);
	//rewrite the data area, the original area won't be cleared
	rewind(gbafp);
	fseek(gbafp,offset2,0);
	for(i=0;i<704;i++)
		fwrite(&New_Bytes[i],sizeof(u8),1,gbafp);
	printf("Finish!");
	fclose(gbafp);
	printf("\nPress Enter to exit...");
	a=getchar();
	return 0;
}
