



* Create a permanent SAS library named ardec;
libname ardec 'W:\ARDEC projects\SAS';

* Read data from Excel files;
*;
proc import out= ardec.coagmet_1999_to_2014
    datafile = 'W:\ARDEC projects\Coagmet_Tidy.xlsx'
    dbms = xlsx replace;
    getnames = yes;
	run;
proc import out= ardec.r1_plant_2006_to_2014
    datafile = 'W:\ARDEC projects\R1_Plant_Tidy.xlsx'
    dbms = xlsx replace;
    getnames = yes;
	run;
proc import out= ardec.r1_soil_2000_to_2014
    datafile = 'W:\ARDEC projects\R1_Soil_Tidy.xlsx'
    dbms = xlsx replace;
    getnames = yes;
	run;
proc import out= ardec.r2_plant_1999_to_2014
    datafile = 'W:\ARDEC projects\R2_Plant_Tidy.xlsx'
    dbms = xlsx replace;
    getnames = yes;
	run;	
proc import out= ardec.r2_soil_1999_to_2014
    datafile = 'W:\ARDEC projects\R2_Soil_Tidy.xlsx'
    dbms = xlsx replace;
    getnames = yes;
	run;
/*data file_list;*/
/*	infile datalines dlm = ',';*/
/*    input filename $ datasetname;*/
/*    datalines;*/
/*	W:\ARDEC projects\R1_Plant_Tidy.xlsx,R1_plant_2006_to_2014*/
/*	W:\ARDEC projects\R1_Soil_Tidy.xlsx,R1_soil_2000_to_2014*/
/*	W:\ARDEC projects\R2_Plant_Tidy.xlsx,R2_plant_1999_to_2014*/
/*	W:\ARDEC projects\R2_Soil_Tidy.xlsx,R2_soil_1999_to_2014*/
/*	W:\ARDEC projects\Coagmet_Tidy.xlsx,Coagmet_1999_to_2014*/
/*   ;*/
/**/
/**/
/*data _null_;*/
/*  set file_list;*/
/*  call execute (*/
/*   "proc import out = " || datasetname || "*/
/*      datafile = '"|| filename ||"'*/
/*      dbms = xlsx replace;*/
/*      *sheet = 'Sheet1';*/
/*      getnames = yes;*/
/*    run;");*/
/*run;*/
