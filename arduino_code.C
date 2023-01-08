#include<stdlib.h>
#include<stdio.h>
#include<string.h>
#include "U8glib.h"
U8GLIB_SH1106_128X64 u8g(U8G_I2C_OPT_NONE);

struct cases{
  float airoffpv;
  int roomTemp;
  int defSp;
  int regSp;
  bool reg_def;
  bool compOn;
  bool evapOn;
  bool condOn;
  bool defOn;
};
bool reg_df;

struct cases *rack=(struct cases*)malloc(sizeof(struct cases));
void regulate(struct cases*);
void show(struct cases*);
void defrost(struct cases*);
void up(struct cases*);
void down(struct cases*);
void defrost(struct cases*);
void serialDisplay(struct cases*);
void sendIOT(struct cases*);

void setup(){
  u8g.begin();
  pinMode(2, INPUT); //Toggle Regulation/Defrost
  pinMode(3, OUTPUT); // Condensor Fan / Evap Fan Command
  pinMode(4, INPUT); // Raise Reg SP
  pinMode(5, OUTPUT); // Compressor Command
  pinMode(6, OUTPUT); //on/off
  pinMode(7, INPUT); // Lower Reg SP 
  pinMode(9, OUTPUT); // reg
  pinMode(10, OUTPUT); // def
  pinMode(11, OUTPUT); // 
  pinMode(13, OUTPUT); // power  
  Serial.begin(9600);
  rack->regSp=20;
  rack->defSp=25;
  bool reg_df=false;
}

void loop(){
  rack->airoffpv=6580.00/analogRead(A0);//8580.00/analogRead(A0);
  rack->roomTemp=analogRead(A1);
  rack->reg_def=digitalRead(2);
  digitalWrite(3, HIGH); //Condensor Fan
  digitalWrite(6,HIGH); // Unit ON Signal
  digitalWrite(13,HIGH);  
  up(rack);
  down(rack);
  if(rack->reg_def==true){
    reg_df=true;
  }
  if(reg_df==true){
    defrost(rack);
  } 
  else{
    regulate(rack);
  } 
  show(rack);
  // serialDisplay(rack);
  sendIOT(rack);
}

void sendIOT(struct cases* rack){
  Serial.write(int(rack->airoffpv));
  // Serial.write(int(rack->regSp));
  // Serial.write(int(rack->defSp));
  // Serial.write(int(rack->compOn));
  // Serial.write(int(reg_df)); 
  delay(2000);  
}

void show(struct cases* obj){
  u8g.firstPage();  
  do{
    u8g.setFont(u8g_font_profont12);
    u8g.setPrintPos(0, 10);
    u8g.print("Arduino IOT Refri");
    u8g.setPrintPos(0, 20);
    u8g.print("Current Temp:");
    u8g.print(rack->airoffpv);
    u8g.setPrintPos(0, 30);
    u8g.print("regulationSP:");
    u8g.print(rack->regSp);
    u8g.setPrintPos(0, 40);
    u8g.print("defrostSP:");
    u8g.print(rack->defSp);
    if(obj->compOn){
       u8g.setPrintPos(0, 50);
       u8g.print("Compressor is On");             
    }
    else{    
       u8g.setPrintPos(0, 50);
       u8g.print("The compressor is Off");
       }
    if(reg_df){
      u8g.setPrintPos(0, 60);
      u8g.print("Defrost is On");
    }
    else{
      u8g.setPrintPos(0, 60);
      u8g.print("Regulation is On");      
    }   
  }while(u8g.nextPage());
}

void regulate(struct cases* obj){
  digitalWrite(9,HIGH);
  digitalWrite(10,LOW);
  if (obj->airoffpv > obj->regSp){
     obj->compOn=true;
     digitalWrite(5, HIGH);
  }
  else if(obj->airoffpv <= obj->regSp){
    obj->compOn=false;
    digitalWrite(5, LOW);    
  }
}

void defrost(struct cases* obj){
  digitalWrite(9,LOW);
  digitalWrite(10,HIGH);
  if (obj->airoffpv > obj->defSp){
     obj->compOn=true;
     digitalWrite(5, HIGH);
     reg_df=false;
  }
  else if(obj->airoffpv <= obj->defSp){
    obj->compOn=false;
    digitalWrite(5, LOW);    
  }
}

void up(struct cases* obj){
  if(digitalRead(4)== true){
    if(reg_df){
      rack->defSp=rack->defSp+1; 
    }
    else{
      rack->regSp=rack->regSp+1;      
    }
        
  }
}

void down(struct cases* obj){
  if(digitalRead(7)== true){
    if(reg_df){
      rack->defSp=rack->defSp-1; 
    }
    else{
      rack->regSp=rack->regSp-1;      
    }
        
  }
}

void serialDisplay(struct cases* obj){
  Serial.print("The Current Temp is : \n");
  Serial.print(obj->airoffpv);
  Serial.print("\n");
  Serial.print("The SP Temp is : \n");
  Serial.print(obj->regSp);
  Serial.print("\n");
  Serial.print("The Defrost Temp is : \n");
  Serial.print(obj->defSp);  
  Serial.print("\n");
  if(reg_df){
      Serial.print("The mode is defrost\n");
      Serial.print("\n");
    }
    else{
      Serial.print("The mode is regulation\n");
      Serial.print("\n");
    }     
}


