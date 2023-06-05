# Oracle Data To Excel
Command line application to extract data from Oracle database and write it to an Excel file.


## Build 

Download the Oracle JDBC driver from [here](https://www.oracle.com/database/technologies/jdbc-drivers-12c-downloads.html) and place it in the `libs` folder.

```
mvn install:install-file -Dfile=libs\ojdbc8.jar -DgroupId=com.oracle -DartifactId=ojdbc8 -Dversion=23.2.0.0.0 -Dpackaging=jar

mvn clean install

mvn dependency:copy-dependencies -DoutputDirectory=target/libs
```

## Run

```
cd target
```
Create a file named `config.properties` in the target folder and add the following content to it.
```
username=your_username
password=your_password
jdbcUrl=your_jdbc_connection_url
```
Create a file named `query.sql` in the target folder with the SQL data extraction query content to it.

Execute the following command to run the application.

```
java -jar OracleDataToExcel-1.0-SNAPSHOT.jar
```