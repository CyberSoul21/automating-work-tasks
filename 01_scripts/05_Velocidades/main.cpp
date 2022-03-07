#include <iostream>
#include <fstream>
#include<cstdlib>
#include<string.h>
#include "IO.h"
#include <conio.h>
/* run this program using the console pauser or add your own getch, system("pause") or input loop */
using namespace std;



string cambio_comilla(string s,int i){//Funcion recursiva para el cambio del caracter comilla (") a espacio 
  
  if(s[i]=='\0'){
    return s;
  }else{
  	if(s[i]=='"'){
	  s[i]=' ';  			    
  	}
    return cambio_comilla(s,i+1);
  };
};

string cambio_comilla(string s){//Funcion recursiva para el cambio del caracter comilla (") a espacio 
	
  return cambio_comilla(s,0);
};

int main(int argc, char** argv) {
	
	bool d=true;//bandera
	int prom=0;//Variavle para el promedio
	int n=0;//Contador de registros
	
	string x;//Lectura de datos
    char p;//Vriable auxiliar
    int ve=0;//Aux
    int *v;//Vector
	
	ifstream archivo("input.txt");//Archivo con datos de entrada, sin titulo
	ofstream salida("output.txt");//Archivo Final
	ofstream modificado("modif.txt");//Archivo auxiliar
	
	do{//Contador de registros
		x = leer_s(archivo);
        if(x!=" "){         
 		 n++;			 		
		}else{
			d=false;
		};
	}while(d);
	d=true;//Cambia la bandera
	
	archivo.close();//Cierra el archivo
	
	ifstream archivo1("input.txt");//Vuelve y lo "crea", para manipularlo
	
	cout<<"Registros: "<<n<<endl;//Imprime en consola el numero total de registros
    v = new int[n];	//Vector para almacenar velocidades
    
   	for(int i=0;i<n;i++){//quita las comillas para poder manipular los datos y manejarlos como enteros
		x = leer_s(archivo1);
		x = cambio_comilla(x);
		escribir_s(modificado,x);
		escribir_s(modificado,"\n");
	}
    modificado.close();//Cierra el archivo
	ifstream archivo2("modif.txt");// y lo crea como archivo de entrada
    
	for(int i=0; i<n; i++){//Clcula el promedio de velocidad
		v[i]= leer_i(archivo2);
		prom+=v[i];		
	}
	prom = prom/n;//operación 
		
	cout<<"El promedio de velocidad es: "<<prom<<" Km/h";//Imprime en consola el promedio 
	
	escribir_s(salida,"El promedio de velocidad es: ");//Guarda el resultado en un archivo plano
	escribir_i(salida,prom);
	escribir_s(salida," Km/h");
	
	archivo2.close();//Cierra los archivos
	salida.close();
	
	return 0;
}
