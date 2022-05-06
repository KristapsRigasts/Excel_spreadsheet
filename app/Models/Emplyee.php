<?php

namespace App\Models;

class Emplyee
{
    private string $name;
    private string $position;
    private array $workingHours;

    public function __construct(string $name, string $position, array $workingHours)
    {
        $this->name = $name;
        $this->position = $position;
        $this->workingHours = $workingHours;
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getPosition(): string
    {
        return $this->position;
    }

    public function getWorkingHours(): array
    {
        return $this->workingHours;
    }
}
