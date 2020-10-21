<?php

namespace Test\Outlaws\Spreadsheet;

use Outlaws\Spreadsheet\ColumnGenerator;
use PHPUnit\Framework\TestCase;

class ColumnGeneratorTest extends TestCase
{
    public function testItGeneratesStartingColumn()
    {
        $generator = new ColumnGenerator();
        $this->assertEquals('A', $generator->getColumn());
    }

    public function testItGeneratesNextColumns()
    {
        $generator = new ColumnGenerator();
        $generator->getColumn();
        $this->assertEquals('B', $generator->getColumn());
        $this->assertEquals('C', $generator->getColumn());
        $this->assertEquals('D', $generator->getColumn());
    }

    public function testItForwardsStartOffset()
    {
        $generator = new ColumnGenerator(null, 10);
        $this->assertEquals('K', $generator->getColumn());
    }

    public function testItForwards()
    {
        $generator = new ColumnGenerator();
        $generator->forward(5);
        $this->assertEquals('F', $generator->getColumn());

        $generator = new ColumnGenerator();
        $generator->getColumn();
        $generator->forward(5);
        $this->assertEquals('G', $generator->getColumn());
    }

    public function testItGeneratesTheCorrectColumnOnZ()
    {
        $generator = new ColumnGenerator(null, 26);
        $this->assertEquals('AA', $generator->getColumn());
    }

    public function testItGeneratesTheCorrectColumnOnZZ()
    {
        $generator = new ColumnGenerator(null, 27*26);
        $this->assertEquals('AAA', $generator->getColumn());
    }

    public function testItForwardsBeyondZ()
    {
        $generator = new ColumnGenerator(null, 30);
        $this->assertEquals('AE', $generator->getColumn());
    }

    public function testItGivesCurrentColumn()
    {
        $generator = new ColumnGenerator();
        $this->assertEquals('A', $generator->getColumn(false));
        $this->assertEquals('A', $generator->getColumn());
        $this->assertEquals('B', $generator->getColumn(false));
    }

    public function testItAddsRowNumber()
    {
        $generator = new ColumnGenerator(5);
        $this->assertEquals('A5', $generator->getColumn());
    }

    public function testItAddsMultiDigitRowNumber()
    {
        $generator = new ColumnGenerator(15);
        $this->assertEquals('A15', $generator->getColumn());
    }

    public function testItResets()
    {
        $generator = new ColumnGenerator();
        $generator->getColumn();
        $generator->getColumn();
        $generator->getColumn();
        $generator->reset();
        $this->assertEquals('A', $generator->getColumn());
    }

    public function testItReturnsCurrentValue()
    {
        $generator = new ColumnGenerator();
        $this->assertEquals('A', $generator->getCurrentColumn());
        $generator->getColumn();
        $this->assertEquals('A', $generator->getCurrentColumn());
        $generator->getColumn();
        $this->assertEquals('B', $generator->getCurrentColumn());
        $this->assertEquals('C', $generator->getColumn(false));
        $this->assertEquals('B', $generator->getCurrentColumn());
    }

    public function testItWalksAllValues()
    {
        $generator = new ColumnGenerator();
        for ($i = 0; $i < 30; $i++){
            $generator->getColumn();
        }

        $letters = [];
        $generator->walk(function($column) use (&$letters) {
            $letters[$column] = 1;
        });

        $this->assertEquals(30, count($letters));
    }

    public function testWalkToZ()
    {
        $generator = new ColumnGenerator();
        $counter = 0;
        $generator->walkTo('Z', function() use (&$counter) {
            $counter++;
        });

        $this->assertEquals(26, $counter);
    }

    public function testItForwardsAndWalksTo()
    {
        $generator = new ColumnGenerator(null, 10);

        $counter = 0;
        $generator->walkTo('Z', function() use (&$counter) {
            $counter++;
        });

        $this->assertEquals(16, $counter);
    }

    public function testItStopsWhenWalkingToUsingLowerCase()
    {
        $generator = new ColumnGenerator();
        $counter = 0;
        // lowercase
        $generator->walkTo('e', function() use (&$counter) {
            $counter++;
        });
        $this->assertEquals(5, $counter);
    }

    public function testItStopsWhenWalkingToLeavingOutRowNumber()
    {
        $generator = new ColumnGenerator(1);
        $counter = 0;
        // lowercase
        $generator->walkTo('e', function() use (&$counter) {
            $counter++;
        });
        $this->assertEquals(5, $counter);
    }
}
