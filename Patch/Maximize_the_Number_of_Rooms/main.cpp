#include <stdio.h>
#include <stdlib.h>         //exit() is in the header

#define u32 unsigned int
#define u8 unsigned char
int main(int argc,char *argv[])
{
	char a;
	if(argc>2)
	{
		printf("Wrong arguments or too many files, you can only draw them in one by one!\nPress any key to Contilue...");
		a=getchar();
		exit(0);
	}
	if(argc==1)
	{
		printf("No files loaded, draw one to the icon of the application!\nPress any key to Contilue...");
		a=getchar();
		exit(0);
	}
	FILE *gbafp;
	if((gbafp=fopen(argv[1],"r+b"))==NULL)
	{
		printf("Cannot Open file!\nPress any key to Contilue...");
		a=getchar();
		exit(0);
	}
	fseek(gbafp,7926400L,0);
	u8 ptrStream[96];
	int i;
	for(i=0;i<96;i++)
		fread(&ptrStream[i],sizeof(u8),1,gbafp);
	fclose(gbafp);
	for(i=0;i<96;i+=4)              //check the value
		printf("%x%x%x%x\n",ptrStream[i],ptrStream[i+1],ptrStream[i+2],ptrStream[i+3]);
	a=getchar();
	return 0;
}