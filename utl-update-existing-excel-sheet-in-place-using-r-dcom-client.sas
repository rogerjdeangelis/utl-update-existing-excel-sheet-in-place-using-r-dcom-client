%let pgm=utl-update-existing-excel-sheet-in-place-using-r-dcom-client;

Update existing excel sheet in place using r dcom client

github
http://tinyurl.com/2p8dyzhv
https://github.com/rogerjdeangelis/utl-update-existing-excel-sheet-in-place-using-r-dcom-client

Where to get RDCOMClient
https://github.com/omegahat/RDCOMClient

Macro dropdown to R on end and in github
https://tinyurl.com/y9nfugth
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

Related repos                                                                                      Related repos
https://github.com/rogerjdeangelis/utl-in-palce-updates-to-an-existing-shared-excel-workbook
https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python
https://github.com/rogerjdeangelis/utl-update-an-excel-workbook-in-place

/*               _     _
 _ __  _ __ ___ | |__ | | ___ _ __ ___
| `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
| |_) | | | (_) | |_) | |  __/ | | | | |
| .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
|_|
*/

/**************************************************************************************************************************/
/*                                                                                                                        */
/* CHANGE THE SEX OF JACK FROM FEMALE TO MALE INPLACE IN AN EXISTING EXCEL SHEET                                          */
/*                                                                                                                        */
/*          INPUT                                  PROCESS                                       OUTPUT                   */
/*          =====                                  =========                                     ========                 */
/* d:/xls/class.xlsx                                                                       d:/xls/class.xlsx              */
/*                                   xlApp <- COMCreate("Excel.Application");                                             */
/*    +-------------------------+    wb<- xlApp[["Workbooks"]]$Open("d:/xls/have.xlsx");      +-------------------------+ */
/*    |     A      |    B       |    sheet <- wb$Worksheets("have");                          |     A      |    B       | */
/*    +-------------------------+      for (roe in c(1:3)) {                                  +-------------------------+ */
/* 1  | NAME       |   SEX      |        cell     <- sheet$Cells(roe,1);                   1  | NAME       |   SEX      | */
/*    +------------+------------+        cellSex  <- sheet$Cells(roe,2);                      +------------+------------+ */
/* 2  | Josh       |    M       |        if (cell[["Value"]] == "Jack" ) {                 2  | Josh       |    M       | */
/*    +------------+------------+              cellSex[["Value"]] <- "M" }                    +------------+------------+ */
/* 3  | Jack       |    F       |      };                                                  3  | Jack       |    M*      | */
/*    +------------+------------+    wb$Save();                                               +------------+------------+ */
/*                                   xlApp$Quit();                                                      *Change Sex to M  */
/*  [HAVE]                                                                                   [HAVE]                       */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

/*----                                                                   ----*/
/*----  CREATE EXCEL WORKBOOK AND SHEET HAVE                             ----*/
/*----                                                                   ----*/

/*----                                                                   ----*/
/*---- Use utl_rbeginx & utl_rendx when replacing mutiple blanks with    ----*
/*---- one blank and  combines lines fails. Somewhat rare, but affects   ----*/
/*---- function read.table                                               ----*/
/*----                                                                   ----*/

/*---- unlink delete the excel file                                      ----*/

%utl_rbeginx;
parmcards4;
library(XLConnect);
unlink("d:/xls/have.xlsx")
have<-read.table(header = TRUE, text = "
NAME SEX
Josh M
Jack F
")
have;
wb <- loadWorkbook("d:/xls/have.xlsx", create = TRUE)
createSheet(wb, name = "have")
writeWorksheet(wb, have, sheet = "have")
saveWorkbook(wb)
;;;;
%utl_rendx;

/**************************************************************************************************************************/
/*                                                                                                                        */
/*          INPUT                                                                                                         */
/*          =====                                                                                                         */
/* d:/xls/class.xlsx                                                                                                      */
/*                                                                                                                        */
/*    +-------------------------+                                                                                         */
/*    |     A      |    B       |                                                                                         */
/*    +-------------------------+                                                                                         */
/* 1  | NAME       |   SEX      |                                                                                         */
/*    +------------+------------+                                                                                         */
/* 2  | Josh       |    M       |                                                                                         */
/*    +------------+------------+                                                                                         */
/* 3  | Jack       |    F       |                                                                                         */
/*    +------------+------------+                                                                                         */
/*                                                                                                                        */
/*  [HAVES]                                                                                                               */
/*                                                                                                                        */
/**************************************************************************************************************************/

 /*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

%utl_submit_r64x('
library(RDCOMClient);
xlApp <- COMCreate("Excel.Application");
wb    <- xlApp[["Workbooks"]]$Open("d:/xls/have.xlsx");
sheet <- wb$Worksheets("have");
  for (roe in c(1:3)) {
    cell     <- sheet$Cells(roe,1);
    cellSex  <- sheet$Cells(roe,2);
    if (cell[["Value"]] == "Jack" ) {
          cellSex[["Value"]] <- "M" }
  };
wb$Save();
xlApp$Quit();
');

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

/**************************************************************************************************************************/
/*                                                                                                                        */
/*          INPUT                                                                                                         */
/*          =====                                                                                                         */
/* d:/xls/class.xlsx                                                                                                      */
/*                                                                                                                        */
/*    +-------------------------+                                                                                         */
/*    |     A      |    B       |                                                                                         */
/*    +-------------------------+                                                                                         */
/* 1  | NAME       |   SEX      |                                                                                         */
/*    +------------+------------+                                                                                         */
/* 2  | Josh       |    M       |                                                                                         */
/*    +------------+------------+                                                                                         */
/* 3  | Jack       |    M       |   Sex xhanged to M                                                                      */
/*    +------------+------------+                                                                                         */
/*                                                                                                                        */
/*  [HAVES]                                                                                                               */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*
 _ __ ___   __ _  ___ _ __ ___  ___
| `_ ` _ \ / _` |/ __| `__/ _ \/ __|
| | | | | | (_| | (__| | | (_) \__ \
|_| |_| |_|\__,_|\___|_|  \___/|___/

*/

%macro utl_rbeginx/des="utl_rbeginx uses parmcards and must end with utl_rendx macro";
%utlfkil(c:/temp/r_pgmx);
%utlfkil(c:/temp/r_pgm);
filename ft15f001 "c:/temp/r_pgm";
%mend utl_rbeginx;

%macro utl_rendx(return=)/des="utl_rbeginx uses parmcards and must end with utl_rendx macro";
run;quit;
* EXECUTE R PROGRAM;
data _null_;
  infile "c:/temp/r_pgm";
  input;
  file "c:/temp/r_pgmx";
  lyn=resolve(_infile_);
  put lyn;
run;quit;
options noxwait noxsync;
filename rut pipe "D:\r412\R\R-4.1.2\bin\r.exe --vanilla --quiet --no-save < c:/temp/r_pgmx";
run;quit;
data _null_;
  file print;
  infile rut;
  input;
  put _infile_;
  putlog _infile_;
run;quit;
data _null_;
  infile " c:/temp/r_pgm";
  input;
  putlog _infile_;
run;quit;
%if "&return" ne ""  %then %do;
  filename clp clipbrd ;
  data _null_;
   infile clp;
   input;
   putlog "xxxxxx  " _infile_;
   call symputx("&return.",_infile_,"G");
  run;quit;
  %end;
%mend utl_rendx;

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
