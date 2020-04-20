#include <iostream>
#include <cstdio>
using namespace std;
int dad[10000];

int hisdad(int d){
    if(dad[d]==d){return d;}

    dad[d] = hisdad(dad[d]);
    return dad[d];
}

int main(){
    int N,M;
    cin>>N>>M;
    for(int i = 1;i <= N;i++){

        dad[i] = i;
    }
    int God,d1,d2;
    
    for(int i = 1;i <= M;i++){
        cin>>God>>d1>>d2;
        if(God == 1){
            
            dad[d1] = hisdad(dad[d2]);
        }else{
            
            if(hisdad(d1) == hisdad(d2)){
                cout<<"Y"<<endl; 
            }else{
                cout<<"N"<<endl; 
            }
        }
    }
    system("pause");
    return 0;
}