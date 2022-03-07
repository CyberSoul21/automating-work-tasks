#ifndef IO_H
#define IO_H
#include <iostream>
 
using namespace std;

double leer(istream& is);
bool leer_b(istream& is);
double leer_i(istream& is);
ostream& escribir(ostream& os,double n);
ostream& escribir_i(ostream& os,int n);
 
ostream& escribir_s(ostream& os,string n);
string leer_s(istream& is);

ostream& escribir_c(ostream& os,char n);
char leer_c(istream& is);

   
#endif // IO_H
