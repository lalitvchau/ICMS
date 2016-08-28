CREATE TABLE  "DOSE" 
   (	"KIDREGNO" NUMBER(10,0), 
	"KIDNAME" VARCHAR2(50), 
	"BCG" DATE, 
	"POLIY" DATE, 
	"POLIY1" DATE, 
	"POLIY2" DATE, 
	"POLIY3" DATE, 
	"DPT1" DATE, 
	"DPT2" DATE, 
	"DPT3" DATE, 
	"HIP1" DATE, 
	"HIP2" DATE, 
	"HIP3" DATE, 
	"VIT1" DATE, 
	"VIT2" DATE, 
	"VIT3" DATE, 
	"VIT4" DATE, 
	"POLIYOBOSTER" DATE, 
	"DPTBOSTER" DATE, 
	"HOPE" NUMBER, 
	"HLP" DATE, 
	"KH9" DATE, 
	"VIT9" DATE, 
	 FOREIGN KEY ("KIDREGNO")
	  REFERENCES  "KIDTABLE" ("KIDREGNO") ENABLE
   )