var fs = require('fs');
var util = require('util');
var xlsx = require('xlsx');
var inquirer = require('inquirer');
var q = require('q');

var baseInsert = 'INSERT INTO %s (%s) VALUES(%s);';
var baseUpdate = 'UPDATE %s SET %s WHERE %s;';
var params = {};

var escapeColname = function(col) {
  return col.split(' ').join('_').toLowerCase();
};

var escapeValue = function(val) {
  return '\'' + val.trim().split('\'').join('\'\'') + '\'';
};

// Write output into file or std output
var writeRows = function(resultRows) {
  var df = q.defer();

  if (params.targetFile) {
    fs.writeFileSync(params.targetFile, resultRows.join('\n'), 'utf8', function(err) {
      if (err) df.reject(err);
    });
  } else {
    resultRows.forEach(function(row) {
      console.log(row);
    });
  }

  df.resolve();
  return df.promise;
}

// Process the xls file, resolve with the SQL statement rows
var processFile = function() {
  var df = q.defer();

  try {
    var wb = xlsx.readFile(params.sourceFile);
  } catch (err) {
    df.reject(err);
  }

  var resultRows = [];
  var sheet = wb.Sheets[wb.SheetNames[0]]; // using 1st sheet only
  var data = xlsx.utils.sheet_to_json(sheet);

  data.forEach(function(row) {
    if (params.method === 'INSERT') {
      var cols = [];
      var vals = [];

      for (var key in row) {
        cols.push(escapeColname(key));
        vals.push(escapeValue(row[key]));
      }

      resultRows.push(util.format(
        baseInsert, params.table, cols.join(', '), vals.join(', ')
      ));
    }
    else if (params.method === 'UPDATE') {
      var setters = [];
      var condition;

      for (var key in row) {
        value = escapeColname(key) + ' = ' + escapeValue(row[key]);
        if (key === params.idColumn) { condition = value; }
        else { setters.push(value); }
      }

      if (!condition) {
        df.reject('PK column not found in row:\n\t' + JSON.stringify(row));
      }

      resultRows.push(util.format(
        baseUpdate, params.table, setters.join(', '), condition
      ));
    }
  });

  df.resolve(resultRows);
  return df.promise;
};

// Starts the prompt, resolve with the answers in a map
var startPrompt = function() {
  return inquirer.prompt([{
      name: 'sourceFile',
      message: 'Source excel file:',
      default: process.argv[2]
    }, {
      name: 'targetFile',
      message: 'Target .sql filename: (leave empty for standard output)'
    }, {
      name: 'table',
      message: 'Target [schema].table?',
      default: 'public.test'
    }, {
      name: 'method',
      message: 'What type of statement to generate?',
      type: 'list',
      choices: ['INSERT', 'UPDATE']
    }, {
      name: 'idColumn',
      message: 'Header title of the Primary Key column?',
      when: function(answers) {
        return answers.method === 'UPDATE';
      },
      validate: function(input) {
        return input && input.length > 0 ? true : 'PK column cannot be empty';
      }
    }
  ]);
};

startPrompt()
  .then(function(answers) {
    params = answers;
    return processFile();
  })
  .then(function(rows) {
    return writeRows(rows);
  })
  .then(function() {
    console.log('All done.')
  })
  .catch(function(err) {
    console.log(err);
    process.exit(1);
  });
