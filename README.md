# Various Scripts

A collection of small, standalone utility scripts for day-to-day admin and
database tasks.

---

## Windows-startup-after-checking-network.bat

A Windows batch script that waits until a specific network/VPN connection is
reachable, then automatically opens a set of everyday websites and
applications — meant to run at Windows startup or login instead of manually
opening the same tabs and apps every morning.

**What it does**

1. **`checkconnection`** — pings the host configured in `servertoping`
   (currently set to the placeholder `hostname` — replace with a real
   internal server/IP that's only reachable once you're on the right
   network/VPN) with a single ping (`ping -n 1`). If the reply contains
   `TTL=` (i.e. a real response came back), it considers the network
   connected.
2. **`startup`** — calls `checkconnection`. If connected, it calls
   `launchwebsites` and `launchapps`. If not, it waits 10 seconds and retries,
   up to **7 attempts** (about 70 seconds total) before giving up.
3. **`launchwebsites`** — opens a fixed list of URLs in the default browser
   (currently intranet/homepage placeholders, a Jira issues view, and
   worldtimebuddy.com).
4. **`launchapps`** — launches a fixed list of desktop applications
   (currently Notepad++ and a placeholder second app).
5. On exit, prints whether startup actions were performed or whether the
   retry limit was reached without ever detecting the network.

**How it's used**

Edit `servertoping`, the URL list in `launchwebsites`, and the executable
paths in `launchapps` to match your own environment, then point a Windows
Startup shortcut (or Task Scheduler "at logon" trigger) at this `.bat` file.
It's designed for a laptop/VPN scenario: Windows starts up faster than the
VPN connects, so the script waits (up to ~70 seconds) until the target host
actually responds before opening anything, instead of firing off apps and
tabs before there's a network connection to use them with.

---

## FindFieldsOnAllTables.sql

A SQL Server query that searches every table in the current database for
columns whose name matches a given pattern, and lists which table (and
schema) each match belongs to.

```sql
SELECT t.name AS table_name, SCHEMA_NAME(schema_id) AS schema_name, c.name AS column_name
FROM sys.tables AS t
INNER JOIN sys.columns c ON t.OBJECT_ID = c.OBJECT_ID
WHERE c.name LIKE '%Created%'
ORDER BY schema_name, table_name;
```

**How it works**

It joins SQL Server's system catalog views `sys.tables` and `sys.columns` on
`OBJECT_ID` to get every column of every table in the database, resolves
each table's schema name with `SCHEMA_NAME(schema_id)`, and filters down to
columns whose name matches the `LIKE` pattern (currently `'%Created%'`, so it
would find columns like `CreatedDate`, `CreatedBy`, `DateCreated`, etc.).
Results are sorted by schema, then table name.

**How it's used**

Run against any SQL Server database when you need to find every table that
has a column matching a naming pattern — useful for things like "which
tables have an audit/creation-date column" without having to check the
schema of every table by hand. Change the string inside the `LIKE` clause to
search for a different column-name pattern.

---

## LastRestoreOfDB.sql

A quick SQL Server query to check when a given database was last restored
from a backup.

```sql
--/### HOW TO CHECK THE LAST REFRESH OF A DATABASE ###/

use MSDB;
go

SELECT MAX(restore_date) as LAST_RESTORE_DT
FROM restorehistory
WHERE destination_database_name = 'DB_NAME'
```

**How it works**

`msdb.dbo.restorehistory` is SQL Server's built-in log of every restore
operation performed on the instance. The query filters that history to a
specific target database (`destination_database_name`) and returns the most
recent (`MAX`) restore date/time for it.

**How it's used**

Replace `'DB_NAME'` with the actual database name you want to check, and run
it against the `msdb` database. Handy for quickly confirming when a
refreshed/restored environment (e.g. a QA or Dev database restored from a
Prod backup) was actually last refreshed, without digging through backup
job history manually.

> Note: `GO` is not a T-SQL keyword — it's a client-side batch separator that
> SSMS/sqlcmd/Azure Data Studio only recognize when it sits alone on its own
> line. It originally shared a line with `USE MSDB`, so no client tool would
> treat it as a separator; it would just be forwarded to the engine as
> literal text and fail with a syntax error. Fixed by putting `GO` on its
> own line (adding the `;` after `USE MSDB` is just good practice, not the
> actual fix — `GO` still has to be alone on its line either way).
