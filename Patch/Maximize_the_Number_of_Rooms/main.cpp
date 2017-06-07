#include <stdio.h>
#include <stdlib.h>         //exit() is in the header

#define u32 unsigned int
#define u8 unsigned char

int main(int argc,char *argv[])
{
	unsigned char a;
	if(argc>2)
	{
		printf("Wrong arguments or too many files, you can only draw them in one by one!\nPress Enter to Contilue...");
		a=getchar();
		exit(0);
	}
	if(argc==1)
	{
		printf("No files loaded, draw one to the icon of the application!\nPress Enter to Contilue...");
		a=getchar();
		exit(0);
	}
	FILE *gbafp;
	if((gbafp=fopen(argv[1],"r+b"))==NULL)
	{
		printf("Cannot Open file!\nPress Enter to Contilue...");
		a=getchar();
		exit(0);
	}
	//First check if this file has been patched by this program
	fseek(gbafp,7928180L,0);
	fread(&a,sizeof(u8),1,gbafp);
	if((unsigned int) a!=255)
	{
		printf("This file has been patched by this program!\nPress Enter to Contilue...");
		a=getchar();
		exit(0);
	}
	//Get all the pointers point to Level Pointers and Flags Stream
	rewind(gbafp);
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
	for(i=0;i<24;i++)
	{
		rewind(gbafp);
		fseek(gbafp,pointers[i],0);
		fread(&B[i],704*sizeof(u8),1,gbafp);
	}
	for(i=0;i<17;i++)                //for The 17 Level and the No.24 can be check somewhere else
	{
		for(j=0;j<17;j++)
		{
			if( ((unsigned int) Bytes[12*j]==i) & ((unsigned int) Bytes[12*j+1]<16) )
				break;
		}
		if(j!=17)             //probably there is no way to judge if the Data areas of Levels have been maximized, is there?
		{
			printf("%2x\n",(unsigned int) Bytes[12*j]);     //temporary output for check
		}
	}
	fclose(gbafp);
	a=getchar();
	return 0;
}
