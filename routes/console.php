<?php

use App\Jobs\CreateExcelWorkSheetJob;

use Illuminate\Support\Facades\Artisan;

/*
|--------------------------------------------------------------------------
| Console Routes
|--------------------------------------------------------------------------
|
| This file is where you may define all of your Closure based console
| commands. Each Closure is bound to a command instance allowing a
| simple approach to interacting with each command's IO methods.
|
*/

Artisan::command('createExcel {monthAndYear}', function ($monthAndYear){
    dispatch(new CreateExcelWorkSheetJob($monthAndYear));
})->purpose('Create excel sheet with employees and working hours by year and month');
