#include <iostream>
#include <fstream>
#include <cstdlib>
#include <string.h>
#include "IO.h"
#include <conio.h>

using namespace std;

string cambiar(string x){//Funcion para cambiar las comillas de por comilla simple
	for(int i=0;x[i] != '\0';i++){
		if(x[i]=='"'){
	    	x[i]=39;	       	
		};	
	};
	return x;
}

int main(int argc, char** argv) {
	/*
	string x = "hola;mundo;puto";
	int c=2,c2=0;
	char* v;
	for(int i=0;x[i] != '\0';i++){
	    c++;
		if(x[i]==';'){
	    	x[i]=',';
	    	c+=2;
	       	
		};	
	};
	v = new char[c];
	v[0]=39;
	c2=1;
	for(int i=0;x[i] != '\0';i++){
			v[c2]=x[i];
			if(v[c2]==','){
				v[c2]=39;
				v[c2+1]=',';
				v[c2+2]=39;	
				c2+=2;			
			};
			c2++;		
	}
	v[c-1]=39;
	cout<<v<<" "<<x<<endl;
	*/
	string x,c2;//variable para almacenar el registro
	int n=0;//Contdor de registros de entrada
	char x1[256];//Variable auxiliar
	bool d=true;//bandera
	char c1[256];
	string txt = ".txt";
	
	ifstream entrada("actualizacion.txt");//Entrada de texto planao importada desde Mapinfo
	ofstream modificado("modif.txt");//Archivo Auxiliar

	
	//entrada.getline(x1,256);
	//x = x1;
	//cout<<x;
	//entrada.getline(x1,256);
	//x = x1;
	//cout<<endl<<x;
	

	
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


	do{ //Ciclo para contar registros de entrada
	entrada.getline(x1,256);//Lee la linea completa con espacios incluidos, se almacena en un vector tipo char
	x = x1;//Se hace asignacion a String	
	//cout<<x<<endl;
        if(x!=""){         
 		 n++;			 		
		}else{
			d=false;
		};
	}while(d);
	d=true;
	entrada.close();//Se cierra archivo de entrada
	ifstream entrada2("actualizacion.txt");//Se vuelve abrir archivo entrada
	cout<<"Registros: "<<n;//Impresion de registros en consola
	
	for(int i=0;i<n;i++){//ciclo para cambiar las comillas (") por comilla simple(')
		entrada2.getline(x1,256);
	    x = x1;		
		x = cambiar(x);
		escribir_s(modificado,x);
		escribir_s(modificado,"\n");
	};
	entrada2.close();
	modificado.close();//Se cierran los archivos
	ifstream entrada3("modif.txt");//Se abre el archivo modificado
	cout<<endl;//Salto de linea en consola
	for(int i=0;i<n;i++){//Ciclo para concatenar la ubicacion del arcchivo con el comando de inserción SQL
	    entrada3.getline(x1,256);
	    x=x1;	    
	    escribir_s(salida,"INSERT INTO `wavetrack`.`tbl_revgeo` (`id_revgeo` ,`name` ,`city` ,`latitude` ,`longitude` ,`geometry` ,`created_by` ,`created_at` ,`created_from` ,`modified_by` ,`modified_at` ,`modified_from` ,`delete_mark`) VALUES (NULL ,"+x+",GeomFromText(CONCAT('POINT(', latitude,' ', longitude,')')),'1',NOW( ),'190.84.246.133','1',NOW(),'190.84.246.133','0');");
		escribir_s(salida,"\n");
		//x=14;
		cout<<x<<endl;//Se imprime en consola todos los archivos
	};	
	entrada3.close();
	salida.close();//El archivo final queda en la carpeta raiz con el nombre "output"
	return 0;
}
