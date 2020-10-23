<?php

namespace Test\Outlaws\Spreadsheet;

use Outlaws\Spreadsheet\ColumnGenerator;
use PHPUnit\Framework\TestCase;

class ColumnGeneratorTest extends TestCase
{
    public function testItGeneratesStartingColumn()
    {
        $generator = $this->getColumnGenerator();
        $this->assertEquals('A', $generator->getColumn());
    }

    public function testItGeneratesNextColumns()
    {
        $generator = $this->getColumnGenerator();
        $generator->getColumn();
        $this->assertEquals('B', $generator->getColumn());
        $this->assertEquals('C', $generator->getColumn());
        $this->assertEquals('D', $generator->getColumn());
    }

    public function testItForwardsStartOffset()
    {
        $generator = $this->getColumnGenerator(null, 10);
        $this->assertEquals('K', $generator->getColumn());
    }

    public function testItForwards()
    {
        $generator = $this->getColumnGenerator();
        $generator->forward(5);
        $this->assertEquals('F', $generator->getColumn());

        $generator = $this->getColumnGenerator();
        $generator->getColumn();
        $generator->forward(5);
        $this->assertEquals('G', $generator->getColumn());
    }

    public function testItGeneratesColumnOnZ()
    {
        $generator = $this->getColumnGenerator(null, 26);
        $this->assertEquals('AA', $generator->getColumn());
    }

    public function testItGeneratesColumnOnZZ()
    {
        $generator = $this->getColumnGenerator(null, 27*26);
        $this->assertEquals('AAA', $generator->getColumn());
    }

    /**
     * 16384 is max columns in Excel
     */
    public function testItGenerates16384()
    {
        $gen = $this->getColumnGenerator();
        for ($i = 0; $i < 16384; $i++) {
            echo $gen->getColumn() . "\n";
        }
        $this->assertEquals('XFD', $gen->getCurrentColumn());
    }

    public function testItForwardsBeyondZ()
    {
        $generator = $this->getColumnGenerator(null, 30);
        $this->assertEquals('AE', $generator->getColumn());
    }

    public function testItGivesCurrentColumn()
    {
        $generator = $this->getColumnGenerator();
        $this->assertEquals('A', $generator->getColumn(false));
        $this->assertEquals('A', $generator->getColumn());
        $this->assertEquals('B', $generator->getColumn(false));
    }

    public function testItAddsRowNumber()
    {
        $generator = $this->getColumnGenerator(5);
        $this->assertEquals('A5', $generator->getColumn());
    }

    public function testItAddsMultiDigitRowNumber()
    {
        $generator = $this->getColumnGenerator(15);
        $this->assertEquals('A15', $generator->getColumn());
    }

    public function testItResets()
    {
        $generator = $this->getColumnGenerator();
        $generator->getColumn();
        $generator->getColumn();
        $generator->getColumn();
        $generator->reset();
        $this->assertEquals('A', $generator->getColumn());
    }

    public function testItReturnsCurrentValue()
    {
        $generator = $this->getColumnGenerator();
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
        $generator = $this->getColumnGenerator();
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
        $generator = $this->getColumnGenerator();
        $counter = 0;
        $generator->walkTo('Z', function() use (&$counter) {
            $counter++;
        });

        $this->assertEquals(26, $counter);
    }

    public function testItForwardsAndWalksTo()
    {
        $generator = $this->getColumnGenerator(null, 10);

        $counter = 0;
        $generator->walkTo('Z', function() use (&$counter) {
            $counter++;
        });

        $this->assertEquals(16, $counter);
    }

    public function testItStopsWhenWalkingToUsingLowerCase()
    {
        $generator = $this->getColumnGenerator();
        $counter = 0;
        // lowercase
        $generator->walkTo('e', function() use (&$counter) {
            $counter++;
        });
        $this->assertEquals(5, $counter);
    }

    public function testItStopsWhenWalkingToLeavingOutRowNumber()
    {
        $generator = $this->getColumnGenerator(1);
        $counter = 0;
        // lowercase
        $generator->walkTo('e', function() use (&$counter) {
            $counter++;
        });
        $this->assertEquals(5, $counter);
    }

    private function getColumnGenerator(?int $rowNumber = null, int $startOffset = 0): ColumnGenerator
    {
        return new ColumnGenerator($rowNumber, $startOffset);
    }
}
