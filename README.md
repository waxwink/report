## Report

this package wraps around the php office spreadsheet to produce excel tables out of laravel collections.

### Getting Started
Install the package with composer :
```
$ composer require waxwink/report
```

### Instruction
Here are some examples to use it :

```php
require __DIR__."/autoload.php";

use Waxwink\Report\Excel;

$keys = [
    'first_att'=> 'First Title',
    'second_att'=> 'Second Title',
    'third_att'=> 'Third Title',
];

$collection = Model::all();

$xl = new Excel($Estate, $keys);

//this would save the file in the root folder
$xl->export("excel_table");

$xl->setColumnWidth('K', 60)
    ->setColumnWidth('L', 60)
    ->wrapTextInColumn('K')
    ->wrapTextInColumn('L')
    
//this would update the file in the root folder 
$xl->update("excel_table");

//this would send the file to the client to download
// works only if you're using laravel
return response()->download($xl->update());

```
