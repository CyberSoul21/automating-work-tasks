#include <iostream>
#include <iostream>
#include <fstream>
#include<cstdlib>
#include<string.h>
#include "IO.h"
#include <conio.h>


/* run this program using the console pauser or add your own getch, system("pause") or input loop */

using namespace std;

int main(int argc, char** argv) {
	cout<<"Desarrollado por: Wilson J. Almario - WaveComm Corporation"<<"\n"<<endl;
	char a1[256];
	char b1[256];//Ingreso de datos
	char c1[256];
	int r,v,n,c=1;//Registros
	string a,b,c2;//Nombres de los puntos A y B
	
	string txt = ".txt";	
	cout<<"Digite punto A: ";
	//getline(cin,a);
	cin.getline(a1,256);
	//cin>>a;
	//cout<<endl<<a1;
	a = a1;
	//cout<<endl<<a;
	cout<<"Digite punto B: ";
	cin.getline(b1,256);
	b = b1;
	cout<<endl;
	cout<<"El Campo Quedaria Asi: Kilometro x "<<a<<" - "<<b;
	cout<<"\n"<<endl;
	
	

	
	cout<<"Nombre del archivo: ";//Creacion del archivo de texto, se solicita nombre 
	cin.getline(c1,256);
	for(int i=0;i<=256;i++){//Se crea directiva .txt
		if(c1[i]=='\0'){
			for(int j=0;j<4;j++){
				c1[i]=txt[j];
				i++;
			}
			c1[i]='\0';//Fin de cadena
			i=256;
		}
	}
	c2=c1;
	cout<<c2;
	ofstream salida(c1);//Crea archivo con los registros finales
	
	
	cout<<"\nIngrese numero de registros: ";
	cin>>r;
	cout<<endl;
	


	
	if(r%2==0)
	{
	  n = r/2;
	  v=n;	
	}else{
	  n = r/2 + 1;
	  v = n-1;
	};
	escribir_s(salida,"ID;NAME\n");
	for(int i=1; i<=n; i++){
		escribir_i(salida,c);
		escribir_s(salida,";Kilometro ");
		escribir_i(salida,i);
		escribir_s(salida, " "+a+" "+"-"+" "+b+"\n");
		c++;		
	};

	//cout<<r;
	for(int i=r;i>n;i--){
		escribir_i(salida,c);
		escribir_s(salida,";Kilometro ");
		escribir_i(salida,v);
		escribir_s(salida, " "+b+" "+"-"+" "+a+"\n");
		c++;
		v--;
	};
	salida.close();
	

	
	return 0;
}
