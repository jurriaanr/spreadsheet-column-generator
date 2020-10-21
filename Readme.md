# Spreadsheet Column Generator

When you have a dynamic number of columns to generate a spreadsheet for it can become a real burden to generate them. Especially if you also want to set styling information on those columns as well, f.e. in Excel.

This library was written with PhpOffice\PhpSpreadsheet (https://github.com/PHPOffice/PhpSpreadsheet) in mind. An excellent library to generate csv or xlsx files.

Each time you call the generator it will generate the next column/cell name. So A-Z and than AA-ZZ etc. Optionally with a row number.

    $columnGenerator = new ColumnGenerator();
    $columnGenerator->getColumn(); // A
    $columnGenerator->getColumn(); // B
    $columnGenerator->getColumn(); // C
    ...
    $columnGenerator->getColumn(); // Z
    $columnGenerator->getColumn(); // AA

A real world example would be this

    $spreadsheet = new Spreadsheet();
    $productsSheet = $spreadsheet->createSheet();
    $productsSheet->setTitle("Products");
    
    $columnGenerator = new ColumnGenerator(1);

    $productsSheet->setCellValue($columnGenerator->getColumn(), 'name'); // A1
    $productsSheet->setCellValue($columnGenerator->getColumn(), 'sku');  // B1
    
    $properties = $this->getBigListOfProperties();
    foreach ($properties as $property) {
        $productsSheet->setCellValue($columnGenerator->getColumn(), $property->getName());
    }

## Walking the previously generated columns

For setting the autosize f.e. on each sell you could of course add another line in the for loop for each cell. Easier though is to walk through all generated cells and execute a callback on each generated column:

    $columnGenerator->walk(function ($columnName) use ($productsSheet) {
        $productsSheet->getColumnDimension(substr($columnName, 0, -1))->setAutoSize(true);
    }); 

It is also possible to do this without first generating cells

    $columnGenerator->walkTo('ZZ1', function ($columnName) use ($productsSheet) {
        ...
    });

You don't have to be to strict on the given column name, this will work as well (lowercase and leaving out row number):
    
    $columnGenerator->walkTo('zz', function ($columnName) use ($productsSheet) {
        ...
    });

    
## Forwarding / Skipping the first _n_ columns

It is possible you already added multiple columns and just want the rest to be dynamic. You can start from a given number of columns:

    $columnGenerator = new ColumnGenerator(null, 10);
    $columnGenerator->getColumn(); // K

It is also possible to forward later on, although the walk methods will still walk these

    $columnGenerator = new ColumnGenerator(null, 10);
    $columnGenerator->getColumn(); // A
    $columnGenerator->getColumn(); // B
    $columnGenerator->forward(5);
    $columnGenerator->getColumn(); // H

## Getting the right value

Not all values are created equal. Depending on what you want to do with the generated column name and when there are different options.
First of all, the getColumn() method has an optional $movePointerForward parameter. This defaults to true. That means that after receiving the value from the generator, internally the pointer will move forward.

    $columnGenerator = new ColumnGenerator();
    $columnGenerator->getColumn();          // A, internally moves to B
    $columnGenerator->getColumn(false);     // B
    $columnGenerator->getColumn(false);     // B
    $columnGenerator->getColumn(true);      // B, internally moves to C
    $columnGenerator->getColumn(false);     // C

So after calling $columnGenerator->getColumn(true) it is not easy to get the value that was returned _again_, since internally the state is already at the next column. For this reason you can call the getCurrentColumn() method, which will return whatever the last call to getColumn(true) returned before moving the pointer. 

    $columnGenerator = new ColumnGenerator();
    $columnGenerator->getCurrentColumn();  // A, this is for convenience, in reality getColumn(true) has not yet been called
    $columnGenerator->getColumn();         // A, internally moves to B
    $columnGenerator->getCurrentColumn();  // A
    $columnGenerator->getColumn(false);    // B
    $columnGenerator->getCurrentColumn();  // A

Some real world use for this could be:

    $spreadsheet = new Spreadsheet();
    $productsSheet = $spreadsheet->createSheet();
    $properties = $this->getBigListOfProperties();
   
    $columnGenerator = new ColumnGenerator(1);
   
    foreach ($properties as $property) {
        $productsSheet->setCellValue($columnGenerator->getColumn(), $property->getName());
    }
    
    // set bold on each generated column
    $productsSheet->getStyle("A1:" . $columnGenerator->getCurrentColumn())->getFont()->setBold(true);

## Reset

You can reset the generator, it will start all over again

    $columnGenerator->getColumn(); // ZX
    $columnGenerator->getColumn(); // ZY
    $columnGenerator->reset();
    $columnGenerator->getColumn(); // A
    
## Forever

Since it is a generator, potentially you can keep going forever. Of course there is a limit to what Excel actually supports.
