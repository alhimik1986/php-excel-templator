<?php

namespace alhimik1986\PhpExcelTemplator\setters\DTO;

class FormulaValue
{

    protected string $formula;

    protected int $quantity;

    public function __construct(string $formula, int $quantity)
    {
        $this->formula = $formula;
        $this->quantity = $quantity;
    }

    public function getFormula(): string
    {
        return $this->formula;
    }

    public function getQuantity(): int
    {
        return $this->quantity;
    }

}