# tika-ml-tools
Using Apache Tika and MarkLogic together

## Getting started

Start by editing **example.config.properties** in the **/src/main/resources** directory:

```properties
host = localhost
username = username
password = password
port = 8000
contentbase = Documents
pstfile = /path/to/inbox.pst
```

In the above example, we're using the default application server on port 8000 (which also allows XCC connections), we're loading all emails into the Documents database and we're using the filename of an example PST file.


## Running the application

You can use gradle (or the wrapper) for this:

```
./gradlew run
``` 