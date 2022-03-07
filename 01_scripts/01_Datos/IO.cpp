#include "IO.h"

double leer(istream& is){
	double a =0;
	is >> a;
	return a;
};
double leer_i(istream& is){
	double a =0;
	is >> a;
	return a;
};
string leer_s(istream& is){
	string a = " ";
	is >> a;
	return a;
};
char leer_c(istream& is){
	char a;
	is >> a;
	return a;
};
bool leer_b(istream& is){

	bool valor = false;
	is >> valor;
	return valor;

};
ostream& escribir_s(ostream& os,string n)
{
	//n=(n)*(-1);
	os << n;
	return os;	
}; 
ostream& escribir_c(ostream& os,char n)
{
	//n=(n)*(-1);
	os << n;
	return os;	
};

ostream& escribir(ostream& os,double n)
{
	//n=(n)*(-1);
	os << n;
	return os;	
}; 
ostream& escribir_i(ostream& os,int n)
{
	//n=(n)*(-1);
	os << n;
	return os;	
};


