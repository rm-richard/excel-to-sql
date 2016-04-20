## excel-to-sql
Node.js script to convert .xls and .xlsx files into INSERT or UPDATE statements.

## Install and run
```sh
$ npm install
$ node index.js
```

## Prompts
  1. **Source excel file**
  2. **Target .sql filename**
  3. **Target [schema].table**
  4. **What type of statement to generate?** INSERT / UPDATE
  5. **Header title of the Primary Key column?** (only if UPDATE was chosen)

## License
MIT
