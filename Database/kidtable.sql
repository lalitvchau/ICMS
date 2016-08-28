CREATE TABLE  "KIDTABLE" 
   (	"COUPLENO" NUMBER(10,0), 
	"KIDREGNO" NUMBER(10,0), 
	"AGGANID" VARCHAR2(20), 
	"KIDNAME" VARCHAR2(30), 
	"MOTHERNAME" VARCHAR2(50), 
	"FATHERNAME" VARCHAR2(50), 
	"BIRTHDATE" DATE, 
	"GENDER" VARCHAR2(10), 
	"KIDWEIGHT" NUMBER(10,0), 
	"PHOTO" VARCHAR2(4000), 
	 PRIMARY KEY ("KIDREGNO") ENABLE, 
	 FOREIGN KEY ("COUPLENO")
	  REFERENCES  "MOTHERTABLE" ("COUPLENO") ENABLE, 
	 FOREIGN KEY ("AGGANID")
	  REFERENCES  "AGGANWARITABLE" ("AGGANID") ENABLE
   )

