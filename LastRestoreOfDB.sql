--/### HOW TO CHECK THE LAST REFRESH OF A DATABASE ###/

use MSDB go

SELECT MAX(restore_date) as LAST_RESTORE_DT FROM restorehistory WHERE destination_database_name = 'DB_NAME'