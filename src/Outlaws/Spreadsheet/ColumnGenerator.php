<?php

declare(strict_types=1);

namespace Outlaws\Spreadsheet;

use Iterator;

class ColumnGenerator
{
    private $generator;
    private $letters;
    private $rowNumber;
    private $startOffset;
    private $previous;

    /**
     * @param int|null $rowNumber the row number to add to the column name, this is optional
     * @param int $startOffset The amount of columns to skip
     */
    public function __construct(?int $rowNumber = null, int $startOffset = 0)
    {
        $this->letters = range('A', 'Z');
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
    public function reset(): self
    {
        $this->init();
        return $this;
    }


    /*
     * For each of the columns generated so far execute the given callable on the column name
     */
    public function walk(callable $callable): self
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
    public function walkTo(string $columnName, callable $callable): self
    {
        $clean = function($s) { return $s ? strtoupper(preg_replace("/\d+/", '', $s)) : $s; };
        $columnName = $clean($columnName);
        $generatedColumnName = null;
        while ($clean($generatedColumnName) !== $columnName) {
            $generatedColumnName = $this->getColumn(true);
            $callable($generatedColumnName);
        }
        return $this;
    }

    public function forward(int $amount): self
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
        $prefix = ''; // the letters
        $index = 0; // the index for the last letter
        $index2 = 0; // the index for the letter before the last letter

        while (true) {
            yield $prefix . $this->letters[$index] . ($rowNr ?? '');
            if ($index < 25) {
                // while there are still letters in the alphabet (so not Z) just continue going up
                $index++;
            } else {
                // reset the last letter on the column name to A
                $index = 0;
                if ($index2 < 26) {
                    // The letter before the last letter is not yet beyond Z, so we increase it and replace the last
                    // letter of the prefix, so AAB becomes AAC f.e.
                    $prefix = substr($prefix, 0, -1) . $this->letters[$index2];
                    $index2++;
                } else {
                    // The letter before the last letter is now at Z so it needs to be reset to A
                    // als we have to walk back all the letters before that to increase (if necessary)
                    // f.e. AZZ will become BAA
                    $index2 = 0;
                    $newPrefix = '';
                    $prefixIndex = strlen($prefix) - 1;

                    while ($prefixIndex >= 0) {
                        $indexOf = $this->getIndexForCharacter($prefix[$prefixIndex]);
                        if ($indexOf < 25) {
                            // everything as expected, the letter before the letter before the last letter is not a Z
                            // so we can just increase it
                            $newPrefix = substr($prefix, 0, $prefixIndex) . $this->letters[$indexOf + 1] . $newPrefix;
                            break;
                        } else {
                            // we know that the letter was a Z so it has to become an A
                            $newPrefix = 'A' . $newPrefix;
                            // Are we at the start of the old prefix yet?
                            if ($prefixIndex > 0) {
                                // we were not yet at the start of the string, so we need to look at the next character
                                $prefixIndex--;
                            } else {
                                // oh my we were actually at the end of the line here..  f.e. ZZZ, which became ZAA above, now wants
                                // to make the last Z an A, but for that we need to create a new prefix that starts at A so ZZZ becomes AAAA
                                $newPrefix = 'A' . $newPrefix;
                                break;
                            }
                        }
                    }

                    // well we walked through the whole prefix ans updated everything we needed to update
                    $prefix = $newPrefix;
                }
            }
        }
    }

    private function getIndexForCharacter($char = ''): int
    {
        return array_search($char, $this->letters) ?: 0;
    }
}
