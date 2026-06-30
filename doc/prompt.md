Could you generate a prompt for genrerating a CMDB in DB, which data model is in the sql file in db/ directry. The cmdb will work as a cli, the data will be stored in excel file, one tab by table.
The interface is composed of multiple cli command, one to load each table. the input of each cli parameter are:
     an input file in excel format, containg the data to import in one table
     an input file containing the map on how to map the 1st parametr  input file to final cmdb table
     an input file about normalisation, which is an excel file which can map value from input file to value to be inserted in db table
As cli invocation, the cli should load the excel data model will be loaded in memory db, make the import in table, and save the output as new excel file with contains the new data, ready to be reuse in new cli call.
the code should be in python, build with object oriented approach, develop as a senior ingenier.
Please ask any question you need to make sure the prompt generated will be ok.

Use the SQL schema in cmdb_2026-06-28T15_49_08.585Z.sql as the CMDB data model.

Build several Python CLI application with a senior-engineer object-oriented design and argparse for the command interface.

Requirements:

Represent the tables and schema in code, with reusable classes for:
schema loading
import mapping
normalization rules
Excel persistence
CLI command dispatch
Load an in-memory SQLite database schema from the SQL model and/or from an Excel workbook with one sheet per table.
Expose one CLI subcommand per table:
load-server
load-localisation
load-application
load-user
load-ip-address
load-vlan
load-os
load-team
Each subcommand must accept:
--data-file: Excel input file containing rows to import for that table
--map-file: Excel input file defining how source columns map to final CMDB columns
--normalization-file: Excel input file defining normalization rules
Behavior:
load the existing Excel data model into an in-memory DB
apply mapping and normalization to the import rows
insert/merge data into the target table
export the full DB state as a new Excel workbook with one sheet per table
produce an output file ready for reuse in a later CLI call
Implementation details:
use Python 3.12+
use argparse for CLI parsing
use pandas + openpyxl for Excel handling
use sqlite3 or SQLAlchemy for the in-memory database
validate mapping and normalization before insert
report clear errors on missing columns, invalid mappings, or normalization mismatches
keep import pipeline reusable and maintainable
each cli command covers one load in table (but each load can update other table)
Schema tables:

Server
Localisation
Application
User
IP address
VLAN
OS
Team
The generated code should be production-minded, object-oriented, and organized as a CLI ETL tool for CMDB Excel imports.
some sample data should be generated to test the correct behavior.

Coudl you simply the cli experience:
     - create separate cli, one for each table server, IP aaddress to avoid using the 1st parametr (load-server)
     - the one cmdb_model.xlsx must be used both for input and output, but when writing the output (mdb_model.xlsx), a backup copy should be made with a time stamp
     - a configuration file providing path to file, directory configuration



in all load_ script, if value is empty in source column, default value should be used for the target_column. If value in source_column start with a star (*), the text (without the star) should be used as inserted value for the traget_column


if a id value (primary key) is provided in the input file, it should be used as value for the row, and not be generated. if a line is imported and the primary key value already exists in the database, an update should be done of value in the database follwing this rule :

    if field exists in both in db and in import file, value from the file is to be updated in db

   if field is empty in import file, value from the db should be kept

 if field is empty in db, value from the imported is to be updated in db

This applies foe all table, and load_* scripts.

add test case to validate this behavior in pytests.