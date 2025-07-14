%let pgm=utl-export-ms-access-table-from-mdb-database-to-sas-table-without-sas-access-using-powershell;

Export ms access table from mdb database to sas table without sas access using powershell

github
https://tinyurl.com/3em6k2z7
https://github.com/rogerjdeangelis/utl-export-ms-access-table-from-mdb-database-to-sas-table-without-sas-access-using-powershell

download and install oledb for access
https://tinyurl.com/2azm47wd
https://www.microsoft.com/en-us/download/details.aspx?id=54920

download simple.mdb (download raw)
https://tinyurl.com/2azm47wd
https://github.com/rogerjdeangelis/utl-export-ms-access-table-from-mdb-database-to-sas-table-without-sas-access-using-powershell/blob/main/simple.mdb

Note I am not using the time limited office365. This works with office 2010.

  CONTENTS

     1 download and install oledb driver
       https://tinyurl.com/28smn6ku
       accessdatabaseengine_X64.exe

     2 download simple.mdb
       https://tinyurl.com/2azm47wd

     3 example using utl_psbegin and utl_psend drop down

     4 related repos

SOAPBOX ON

Best without the time limited office 365 product.
Best with the classic editor because of dm commands

The mdb format has not changed since the accdb was introduced in 2007.

To be compatible to R and Python, which do not have a interfaces to accdb, it makes sense to use mdb databases.
We don't have to worry about future failures when using the mdb format.

There seems to be an isuue with accdb databases drivers, mdb seems more universal with one driver?

Updates to accdb format

Access 2007 (ACE 12)
Access 2010 (ACE 14)
Access 2013 (ACE 15)
Access 2016 (ACE 16)
Access 2019 (ACE 16)
Access 2021 (ACE 16)

The connection string is really a nasty statement.
Has to be on one line.
Not easy to substitute powershell variables for arguments like database.

SOAPBOX OFF

/**************************************************************************************************************************/
/* INPUT           |PROCESS                                                                              | OUTPUT         */
/* =====           |=======                                                                              | =======        */
/* downloader      |proc datasets lib=work nodetails nolist;                                             | NAME   SEX AGE */
/*                 |delete mdbclass;                                                                     |                */
/*d:\mdb\simple.mdb|run;quit;                                                                            | Alfred  M  14  */
/*                 |                                                                                     | Alice   F  13  */
/* Table have      |%symdel mdb csv table / nowarn;                                                      | Barbara F  13  */
/*                 |                                                                                     | Carol   F  14  */
/* NAME    SEX AGE |%utlfkil(d:\csv\class.csv);                                                          | Henry   M  14  */
/*                 |                                                                                     | James   M  12  */
/* Alfred  M   14  |%utl_psbegin;                                                                        | ....           */
/* Alice   F   13  |parmcards4;                                                                          |                */
/* Barbara F   13  |$mdb="d:\mdb\simple.mdb";                                                            |                */
/* Carol   F   14  |                                                                                     |                */
/* Henry   M   14  |$ConnectionString="Provider=Microsoft.ACE.OLEDB.16.0;Data Source=d:\mdb\simple.mdb;" |                */
/* James   M   12  |                                                                                     |                */
/* ...             |$SQLquery = "SELECT TOP 10 * FROM have"                                              |                */
/*                 |                                                                                     |                */
/*                 |# Create and open the OLE DB connection                                              |                */
/*                 |$conn = New-Object System.Data.OleDb.OleDbConnection                                 |                */
/*                 |$conn.ConnectionString = $ConnectionString                                           |                */
/*                 |$conn.Open()                                                                         |                */
/*                 |                                                                                     |                */
/*                 |# Create and execute the command                                                     |                */
/*                 |$comm = New-Object System.Data.OleDb.OleDbCommand($SQLquery, $conn)                  |                */
/*                 |$adapter = New-Object System.Data.OleDb.OleDbDataAdapter $comm                       |                */
/*                 |$dataset = New-Object System.Data.DataSet                                            |                */
/*                 |$adapter.Fill($dataset) | Out-Null                                                   |                */
/*                 |                                                                                     |                */
/*                 |# Close the connection                                                               |                */
/*                 |$conn.Close()                                                                        |                */
/*                 |                                                                                     |                */
/*                 |# Export the results to CSV                                                          |                */
/*                 |$dataset.Tables[0] | Export-Csv "d:/csv/mdbcsv.csv" -NoTypeInformation               |                */
/*                 |;;;;                                                                                 |                */
/*                 |%utl_psend;                                                                          |                */
/*                 |                                                                                     |                */
/*                 |dm "";                                                                               |                */
/*                 |dm "dimport 'd:/csv/mdbcsv.csv' mdbhave replace";                                    |                */
/*                 |                                                                                     |                */
/*                 |proc print data=mdbhave;                                                             |                */
/*                 |run;quit;                                                                            |                */
/**************************************************************************************************************************/

/*        _       _           _
 _ __ ___| | __ _| |_ ___  __| |  _ __ ___ _ __   ___  ___
| `__/ _ \ |/ _` | __/ _ \/ _` | | `__/ _ \ `_ \ / _ \/ __|
| | |  __/ | (_| | ||  __/ (_| | | | |  __/ |_) | (_) \__ \
|_|  \___|_|\__,_|\__\___|\__,_| |_|  \___| .__/ \___/|___/
                                          |_|
*/
REPO
------------------------------------------------------------------------------------------------------------------------------------
https://github.com/rogerjdeangelis/utl-Importing-data-from-MS-Access-with-Table-names-over-32-characters
https://github.com/rogerjdeangelis/utl-import-dbf-dif-ods-xlsx-spss-json-stata-csv-html-xml-tsv-files-without-sas-access-products
https://github.com/rogerjdeangelis/utl-indirect-addressing-to-access-variable-names
https://github.com/rogerjdeangelis/utl-sas-access-to-the-universe-of-data
https://github.com/rogerjdeangelis/utl-sas-to-and-from-sqllite-excel-ms-access-spss-stata-using-r-packages-without-sas
https://github.com/rogerjdeangelis/utl-unix-or-windows-export-dataset-to-ms-access-mdb-without-sas-access
https://github.com/rogerjdeangelis/utl-without-ms-access-send-sas-dataset-to-access-subset-and-return-table-to-sas-rodbc
https://github.com/rogerjdeangelis/utl_converting_32bit_ms_access_tables_to_sas_datasets_without_sas_access_products
https://github.com/rogerjdeangelis/utl_creating_sas7bdat_from_32bit_or_64bit_ms-access_table_using_wps_express_proc_R
https://github.com/rogerjdeangelis/utl_exporting_longtext_fields_to_ms_access
https://github.com/rogerjdeangelis/utl_importing_long_strings_from_ms_access


/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
