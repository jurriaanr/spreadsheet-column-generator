<?php

declare(strict_types=1);

namespace Outlaws\Spreadsheet;

use Iterator;

class ColumnGenerator implements ColumnGeneratorInterface
{
    private $generator;
    private $rowNumber;
    private $startOffset;
    private $previous;

    /**
     * @param int|null $rowNumber the row number to add to the column name, this is optional
     * @param int $startOffset The amount of columns to skip
     */
    public function __construct(?int $rowNumber = null, int $startOffset = 0)
    {
        $this->rowNumber = $rowNumber;
        $this->startOffset = $startOffset;
        $this->init();
    }

    /*
     * Get the value of the current column. If the $movePointerForward parameter is true, the index will be forwarded
     */
    public function getColumn(bool $movePointerForward = true): string
    {
        $value = (string)$this->generator->current();
        if ($movePointerForward) {
            $this->previous = $value;
            $this->generator->next();
        }
        return $value;
    }

    /*
     * This will be the same as getColumn() before getColumn(true)
     * is called. However, after one of those has been called the internal pointer has moved forward and
     * getColumn() will report the upcoming. To know what value was last returned by the generator before moving the
     * pointer, use this method.
     */
    public function getCurrentColumn(): string
    {
        return $this->previous ?? $this->getColumn(false);
    }

    /*
     * Reset the generator and start over fresh
     */
    public function reset(): ColumnGeneratorInterface
    {
        $this->init();
        return $this;
    }


    /*
     * For each of the columns generated so far execute the given callable on the column name
     */
    public function walk(callable $callable): ColumnGeneratorInterface
    {
        $currentValue = $this->getCurrentColumn();
        if ($currentValue !== null) {
            $this->reset();
            $this->walkTo($currentValue, $callable);
        }
        return $this;
    }

    /*
     * For each of the columns generated up to and including the given column name execute
     * the given callable on the column name.
     * Be careful, the wrong column name may make this loop forever
     * This function will not rewind and will forward the index until it's done.
     * You are a bit free in the columname in that is does not have to be uppercase and numbers are stripped
     */
    public function walkTo(string $columnName, callable $callable): ColumnGeneratorInterface
    {
        $clean = function ($s) {
            return $s ? strtoupper(preg_replace("/\d+/", '', $s)) : $s;
        };
        $columnName = $clean($columnName);
        $generatedColumnName = null;
        while ($clean($generatedColumnName) !== $columnName) {
            $generatedColumnName = $this->getColumn(true);
            $callable($generatedColumnName);
        }
        return $this;
    }

    public function forward(int $amount): ColumnGeneratorInterface
    {
        // the poor mens way of forwarding. The values can probably be calculated as well 🤷‍
        while ($amount-- > 0) {
            $this->generator->next();
        }

        return $this;
    }

    private function init(): void
    {
        $this->generator = $this->getColumnGenerator($this->rowNumber);
        if ($this->startOffset) {
            $this->forward($this->startOffset);
        }
        $this->previous = null;
    }

    private function getColumnGenerator(?int $rowNr = null): Iterator
    {
        $letters = range('A', 'Z');
        $index = [0]; // the indexes for the letters
        $indexLength = 0;

        while (true) {
            $s = '';
            for ($i = 0; $i <= $indexLength; $i++) {
                $s .= $letters[$index[$i]];
            }
            yield $s . ($rowNr ?? '');
            // unfortunately this is 3x as slow
            //yield join('', array_map(function (int $i) use ($letters) { return $letters[$i]; }, $index))  . ($rowNr ?? '');

            // better to not start a for loop if we dont need to
            if ($index[$indexLength] < 25) {
                $index[$indexLength]++;
            } else {
                for ($i = $indexLength; $i >= 0; $i--) {
                    if ($index[$i] < 25) {
                        $index[$i]++;
                        break;
                    } else if ($i > 0) {
                        $index[$i] = 0;
                        $index[$i - 1]++;
                        if ($index[$i - 1] <= 25) {
                            break;
                        }
                    } else {
                        $index[$i] = 0;
                        array_unshift($index, 0);
                        $indexLength++;
                    }
                }
            }
        }
    }
}
