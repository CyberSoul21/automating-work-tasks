#include <iostream>
#include <fstream>
#include<cstdlib>
#include<string.h>
#include "IO.h"
#include <conio.h>
/* run this program using the console pauser or add your own getch, system("pause") or input loop */
using namespace std;
int longitud_cadena(char* s,int i){
  if(s[i]=='\0'){
    return i;
  }else{
    return longitud_cadena(s,i+1);
  };
};
string longitud_cadena_c(string s,int i){
  
  if(s[i]=='\0'){
    return s;
  }else{
  	if(s[i]==')'){
  		if(s[i+1]==','){
  			if(s[i+2]=='('){
  				s[i+1]='\n';
		  	}
		  }
  				    
  	}
  	
    return longitud_cadena_c(s,i+1);
  };
};

string longitud_cadena_c(string s){
	
  return longitud_cadena_c(s,0);
};


string longitud_cadena_al(string s,int i){
  
  if(s[i]=='\0'){
    return s;
  }else{
  	if(s[i]==')'){
  				  s[i]=')';  
  	}
  	
    return longitud_cadena_al(s,i+1);
  };
};

string longitud_cadena_al(string s){
	
  return longitud_cadena_al(s,0);
};





int longitud_cadena(char* s){
  return longitud_cadena(s,0);
};

int main(int argc, char** argv) {
	
	string x;
    char p;
    int ve=0;
    int *v;
   
    
	ifstream archivo("Alejo.txt");
	ofstream salida("output.txt");
	ofstream modificado("modif.txt");
	//ifstream archivo("input2.txt");
	/*
	for(int i=0; i<8251; i++)
	{
		for(int j=0; j<2; j++)
		{
				x = leer_s(archivo);
				x = longitud_cadena_c(x);
	    		escribir_s(salida,x);
	    		escribir_s(salida," ");
	 	};
	 	escribir_s(salida,"\n");
	};
	/*
	/*cout<<"Digite palabra"<<endl;
	cin>>p;
	cout<<p<<endl;
	cout<<longitud_cadena(p)<<endl;
	//strlen();
	*/
	bool d=true;
	int prom=0;
	int n=0;
	//cout<<"Digite numero de registros: "<<endl;
	//cin>>n;
	
	do{
		x = leer_s(archivo);
        if(x!=" "){         
 		 n++;			 		
		}else{
			d=false;
		};
	}while(d);
	d=true;
	archivo.close();
    ifstream archivo1("Alejo.txt");
	cout<<"Registros: "<<n;
    v = new int[n];
	for(int i=0;i<n;i++){
		x = leer_s(archivo1);
		x = longitud_cadena_c(x);
		escribir_s(modificado,x);
		escribir_s(modificado,"\n");
	}
	modificado.close();
	ifstream archivo2("modif.txt");
	//for(int i=0; i<n; i++){
	//	v[i]= leer_i(archivo2);
	//	prom+=v[i];		
	//}
	//prom = prom/n;	
	//cout<<prom;
	//escribir_s(salida,"El promedio de velocidad es: ");
	
	//escribir_i(salida,prom);
	//escribir_s(salida," Km/h");
	
	do{
		x = leer_s(archivo2);
        if(x!=" "){         
 		 n++;			 		
		}else{
			d=false;
		};
	}while(d);
	archivo2.close();
    ifstream archivo3("modif.txt");
    cout<<"\n";
	cout<<"Registros: "<<n;
    v = new int[n];
    for(int i=0;i<n-1;i++){
		x = leer_s(archivo3);
		x = longitud_cadena_al(x);
		escribir_s(salida,x);
		escribir_s(salida,",");
		escribir_s(salida,"\n");
	}
	
		
	archivo1.close();
	archivo3.close();
	salida.close();
		
	
	
	return 0;
}
