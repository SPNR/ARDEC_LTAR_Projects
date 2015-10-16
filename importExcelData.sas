* STILL IN PROGRESS!
* Macro to read data tables from Excel files and save them in library 'ardec';
%macro read_xlsx(filename, dataname);
%put _all_;
	%let path = W:\ARDEC projects\;
	%let ext = .xlsx;
	%let full_path = cat(&path, &filename, &ext);
	%let destination = cat('ardec.', &dataname);
	proc import out = destination;
    	datafile = &full_path;
    	dbms = xlsx replace;
		sheet = 'Sheet1';
    	getnames = yes;
	run;
%mend read_xlsx;
%read_xlsx(R1_Plant_Tidy, r1_plant_2006_to_2014);

* Insert the list of filenames into a single macro variable;
%let filename_list = R1_Plant_Tidy R1_Soil_Tidy;




* Create a permanent SAS library named ardec;
libname ardec 'W:\ARDEC projects\SAS';

* Read data from Excel files;
*;
proc import out= ardec.coagmet_1999_to_2014
    datafile = 'W:\ARDEC projects\coagmetTidy.xlsx'
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

