#include<bits/stdc++.h>
#include<random>
#include<ctime>

using namespace std;

struct data
{
	int start_h=6,off_h=16;
	int start_m=0,off_m=30;
}dat[40][35]; //start work:05:55-06:30;off work:16:30-17:05

bool repeat_start(int i,int j)//i-m,j-n
{
	if(dat[i][j].start_m==dat[i][j-1].start_m) return 1;
	for(int m=1;m<i;m++)
	{
		if(dat[i][j].start_m==dat[m][j].start_m) return 1;
	}
	return 0;
}
bool repeat_off(int i,int j)//i-m,j-n
{
	if(dat[i][j].off_m==dat[i][j-1].off_m) return 1;
	for(int m=1;m<i;m++)
	{
		if(dat[i][j].off_m==dat[m][j].off_m) return 1;
	}
	return 0;
}

int main()
{
	freopen("randtime.json","w",stdout);
	srand(time(0));
	for(int i=1;i<=35;i++)
	{
		for(int j=1;j<=30;j++)
		{
			do
				dat[i][j].start_m=rand()%36;
			while(repeat_start(i,j));
			dat[i][j].start_m-=5;
			if(dat[i][j].start_m<0)
			{
				dat[i][j].start_h--;
				dat[i][j].start_m+=60;
			}
			do
				dat[i][j].off_m=rand()%36;
			while(repeat_off(i,j));
			dat[i][j].off_m+=30;
			if(dat[i][j].off_m>=60)
			{
				dat[i][j].off_h++;
				dat[i][j].off_m-=60;
			}
		}
	}
	cout<<"[";
	for(int i=1;i<=35;i++)
	{
		cout<<"[";
		for(int j=1;j<=30;j++)
		{
			printf("[\"%d:%d\",\"%d:%d\"]",dat[i][j].start_h,dat[i][j].start_m,dat[i][j].off_h,dat[i][j].off_m);
			if(j!=30) cout<<",";
		}
		cout<<"]";
		if(i!=35) cout<<",";
	}
	cout<<"]";
	return 0;
}
