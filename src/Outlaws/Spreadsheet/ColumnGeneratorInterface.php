<?php

namespace Outlaws\Spreadsheet;

interface ColumnGeneratorInterface
{
    public function getColumn(bool $movePointerForward = true): string;

    public function getCurrentColumn(): string;

    public function reset(): self;

    public function walk(callable $callable): self;

    public function walkTo(string $columnName, callable $callable): self;

    public function forward(int $amount): self;
}
