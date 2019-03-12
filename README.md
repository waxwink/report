## Report

this package wraps around the php office spreadsheet to produce excel tables out of laravel collections.

### Getting Started
Install the package with composer :
```
$ composer require waxwink/report
```

### Instruction
Here is an examples to use the package :

```php
use Illuminate\Support\Collection;
use Waxwink\Report\Excel;

require __DIR__."/vendor/autoload.php";

$keys = [
    'id'=> 'Product ID',
    'name'=> 'Name',
    'price'=> 'Price',
];

$collection = new Collection([
    [
        'id'=> '1574',
        'name'=> 'Phone',
        'price'=> '100',],
    [
        'id'=> '6541',
        'name'=> 'Printer',
        'price'=> '150',
    ],
    [
        'id'=> '9652',
        'name'=> 'Laptop',
        'price'=> '350',
    ],
    [
        'id'=> '6971',
        'name'=> 'Mouse',
        'price'=> '30',
    ]
]);

$xl = new Excel($collection, $keys);

//this would save the file in the root folder : table.xlsx
$xl->export('table');

```

you can also send the file to the client to download. :
```php
// (works only if you're using laravel)
return response()->download($xl->update());
```
