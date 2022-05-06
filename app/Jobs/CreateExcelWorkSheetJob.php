<?php

namespace App\Jobs;

use App\Models\Emplyee;
use Carbon\Carbon;
use Carbon\CarbonPeriod;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldBeUnique;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class CreateExcelWorkSheetJob implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    private string $monthAndYear;
    private array $monthsInYear;

    public function __construct(string $monthAndYear)
    {
        $this->monthAndYear = $monthAndYear;
        $this->monthsInYear = [
            '01' => 'janvāris',
            '02' => 'februāris',
            '03' => 'marts',
            '04' => 'aprīlis',
            '05' => 'maijs',
            '06' => 'jūnijs',
            '07' => 'jūlijs',
            '08' => 'augusts',
            '09' => 'septembris',
            '10' => 'oktobris',
            '11' => 'novembris',
            '12' => 'decembris'
        ];
    }

    public function handle()
    {
        [$month, $year] = explode('-',$this->monthAndYear);

        $dt = Carbon::createFromDate($year, $month);
        $endDate = $dt->daysInMonth;

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Set header

        $sheet->mergeCells('C2:AD2')->setCellValue('C2', 'DIENESTA PIENĀKUMU IZPILDE (DARBA) LAIKA UZSKAITES TABULA')
            ->getStyle('C2')->getFont()->setSize(15)->setBold(true);
        $sheet->getStyle('C2')->getAlignment()->setHorizontal('center')->setVertical('center');

        $sheet->mergeCells('C3:AD3')->setCellValue('C3', "par $year. gada {$this->monthsInYear[$month]}")
            ->getStyle('C3')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle('C3')->getFont()->setSize(15)->setBold(true);

        //Set Column names

        $sheet->mergeCells('A5:A8')->setCellValue('A5', 'nr.p/k')
            ->getStyle('A5')->getAlignment()->setTextRotation(90)->setVertical('center');
        $sheet->getColumnDimension('A')->setWidth(3);

        $sheet->mergeCells('C5:AG7')->setCellValue('C5','Meneša datumi')
            ->getStyle('C5')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle('C5')->getFont()->setSize(13)->setBold(true);

        $sheet->mergeCells('B5:B8')->setCellValue('B5', 'vārds, uzvārds, amats')
            ->getStyle('B5')->getAlignment()->setVertical('center')->setHorizontal('center')->setWrapText(true);
        $sheet->getColumnDimension('B')->setWidth(16);

        // Set month days

        $period = CarbonPeriod::create("$year-$month-01", "$year-$month-$endDate");

        foreach ($period as $date) {
            $date->format('Y-m-d');
        }

        $dates = $period->toArray();

        $dateColumnStartingPoint = 3;
        foreach($dates as $index => $date)
        {
            $sheet->setCellValueByColumnAndRow($dateColumnStartingPoint, 8, $index +1);

            if($date->isWeekend())
            {
                $sheet->getStyleByColumnAndRow($dateColumnStartingPoint, 8)
                    ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('ADD8E6');
            }
            $dateColumnStartingPoint ++;
        }

        $monthDateRangeColumns = range(3, 33);
        foreach ($monthDateRangeColumns as $value)
        {
            $sheet->getColumnDimensionByColumn($value)->setWidth(4);
        }

        $sheet->getStyle('C8:AG8')->getAlignment()->setHorizontal('center');

        // Set extra column names

        $sheet->mergeCells('AH5:AH8')->setCellValue("AH5",'stundu skaits mēneša normālajā dien.pienākumu izpildes (darba) laikā')
            ->getStyle('AH5')->getAlignment()->setTextRotation(90)->setWrapText(true);
        $sheet->getRowDimension('8')->setRowHeight(210);
        $sheet->getColumnDimension('AH')->setWidth(6);

        $sheet->mergeCells('AI5:AW5')->setCellValue('AI5','Nostrādātās stundas')
            ->getStyle('AI5')->getAlignment()->setVertical('center')->setHorizontal('center');
        $sheet->getStyle('AI5')->getFont()->setSize(12)->setBold(true);

        $sheet->mergeCells('AI6:AI8')->setCellValue("AI6",'pavisam')
            ->getStyle('AI6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getStyle('AI6')->getFont()->setBold(true);
        $sheet->getColumnDimension('AI')->setWidth(4);

        $sheet->mergeCells('AJ6:AJ8')->setCellValue("AJ6",'naktī (no plkst.22.00 līdz 6.00')
            ->getStyle('AJ6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AJ')->setWidth(4);

        $sheet->mergeCells('AK6:AK8')->setCellValue("AK6",'dežūras ārpus dienesta pienākumu izpildes vietas1')
            ->getStyle('AK6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AK')->setWidth(4);

        $sheet->mergeCells('AL6:AL8')->setCellValue("AL6",'virsstundas')
            ->getStyle('AL6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AL')->setWidth(4);

        $sheet->mergeCells('AM6:AM8')->setCellValue("AM6",'virs noteiktā dienesta pienākumu izpildes laika2')
            ->getStyle('AM6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AM')->setWidth(4);

        $sheet->mergeCells('AN6:AN8')->setCellValue("AN6",'virs noteiktā dienesta pienākumu izpildes laika3')
            ->getStyle('AN6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AN')->setWidth(4);

        $sheet->mergeCells('AO6:AO8')->setCellValue("AO6",'virs noteiktā dienesta pienākumu izpildes laika4')
            ->getStyle('AO6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AO')->setWidth(4);

        $sheet->mergeCells('AP6:AP8')->setCellValue("AP6",'virs noteiktā dienesta pienākumu izpildes laika5')
            ->getStyle('AP6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AP')->setWidth(4);

        $sheet->mergeCells('AQ6:AQ8')->setCellValue("AQ6",'virs noteiktā dienesta pienākumu izpildes laika6')
            ->getStyle('AQ6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AQ')->setWidth(4);

        $sheet->mergeCells('AR6:AR8')->setCellValue("AR6",'svētku dienā s(darbiniekiem)')
            ->getStyle('AR6')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AR')->setWidth(4);

        $sheet->mergeCells('AS6:AW7')->setCellValue("AS6",'virs noteiktā dienesta pienākumu izpildes laika (virsstundas)')
            ->getStyle('AS6')->getAlignment()->setHorizontal('center')->setWrapText(true);
        $sheet->getStyle('AS6')->getFont()->setSize(9);
        $sheet->getRowDimension('6')->setRowHeight(25);

        $sheet->setCellValue("AS8",'kopā iepriekšējo 4 mēnešu periodā')
            ->getStyle('AS8')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AS')->setWidth(4);

        $sheet->setCellValue("AT8",'pirmajā mēnesī (janvāris)')
            ->getStyle('AT8')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AT')->setWidth(4);

        $sheet->setCellValue("AU8",'pirmajā mēnesī (februāris)')
            ->getStyle('AU8')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AU')->setWidth(4);

        $sheet->setCellValue("AV8",'pirmajā mēnesī (marts)')
            ->getStyle('AV8')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AV')->setWidth(4);

        $sheet->setCellValue("AW8",'pirmajā mēnesī (aprīlis)')
            ->getStyle('AW8')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AW')->setWidth(4);

        $sheet->mergeCells('AX5:AX8')->setCellValue("AX5",'virs noteiktā dienesta pienāk. izp. laika kopā četru menešu periodā')
            ->getStyle('AX5')->getAlignment()->setTextRotation(90)->setHorizontal('center')->setWrapText(true);
        $sheet->getColumnDimension('AX')->setWidth(4);

        $sheet->mergeCells('AY5:AY8')->setCellValue("AY5",'darba dienas')
            ->getStyle('AY5')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AY')->setWidth(4);
        $sheet->getStyle('AY5')->getFont()->setBold(true);

        $sheet->mergeCells('AZ5:AZ8')->setCellValue("AZ5",'atvaļinājuma dienas')
            ->getStyle('AZ5')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('AZ')->setWidth(4);
        $sheet->getStyle('AZ5')->getFont()->setBold(true);

        $sheet->mergeCells('BA5:BA8')->setCellValue("BA5",'darbanespējas dienas')
            ->getStyle('BA5')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('BA')->setWidth(4);
        $sheet->getStyle('BA5')->getFont()->setBold(true);

        $sheet->mergeCells('BB5:BB8')->setCellValue("BB5",'apmaksāts atpūtas laiks')
            ->getStyle('BB5')->getAlignment()->setTextRotation(90)->setHorizontal('center');
        $sheet->getColumnDimension('BB')->setWidth(4);
        $sheet->getStyle('BB5')->getFont()->setBold(true);

        $sheet->getStyle('A5:BB8')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);

        //Set gathered info about employees

        $employees = [
            new Emplyee('Māris', 'Biroja priekšnieks',[5,8,9,'B','B',9,7,5,6]),
            new Emplyee('Jānis', 'vecākais inspektors',[5,8,9,'B',7,9,7,5,6,9,7,5,6]),
            new Emplyee('Gatis', 'Operatīvais darbinieks',[5,8,9,'B',7,5,6,9,7,5,6]),
            new Emplyee('Sandis', 'Operatīvais dežurants',[5,8,9,'B',7,5,6,9,'B',5,6])
        ];

        $employeeColumnStartingPoint = 9;
        $employeeColumnEndingPoint = 11;

        foreach($employees as $index => $employee)
        {
            $sheet->mergeCellsByColumnAndRow(1, $employeeColumnStartingPoint, 1, $employeeColumnEndingPoint)
                ->setCellValueByColumnAndRow(1, $employeeColumnStartingPoint, $index + 1)
                ->getStyleByColumnAndRow(1, $employeeColumnStartingPoint)->getAlignment()
                ->setVertical('center')->setHorizontal('center');

            $nameAndPosition = new RichText();
            $nameAndPosition->createTextRun($employee->getName())->getFont()->setBold(true);
            $nameAndPosition->createText(" {$employee->getPosition()}");

            $sheet->setCellValueByColumnAndRow(2,$employeeColumnStartingPoint,$nameAndPosition);
            $sheet->getStyleByColumnAndRow(2,$employeeColumnStartingPoint)->getAlignment()->setWrapText(true);

            $sheet->setCellValueByColumnAndRow(2,$employeeColumnStartingPoint +1, 't.sk. 22.00-06.00');

            $sheet->setCellValueByColumnAndRow(2,$employeeColumnStartingPoint +2, 'izsaukums (plkst. no līdz)');
            $sheet->getStyleByColumnAndRow(2,$employeeColumnEndingPoint)->getFont()->setSize(7);

            $sheet->getStyleByColumnAndRow(1,$employeeColumnStartingPoint, 54, $employeeColumnEndingPoint)
                ->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);

            $sheet->getStyleByColumnAndRow(1,$employeeColumnStartingPoint, 54, $employeeColumnEndingPoint)
                ->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THICK);

            $sheet->getStyleByColumnAndRow(3,$employeeColumnStartingPoint, 33, $employeeColumnEndingPoint)
                ->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);

            $sheet->getStyleByColumnAndRow(3,$employeeColumnStartingPoint, 33, $employeeColumnEndingPoint)
                ->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);

            $hourColumn =3;
            foreach ($employee->getWorkingHours() as $hours)
            {
                $sheet->setCellValueByColumnAndRow($hourColumn, $employeeColumnStartingPoint, $hours)
                    ->getStyleByColumnAndRow($hourColumn,$employeeColumnStartingPoint)->getAlignment()
                    ->setVertical('center')->setHorizontal('center');
                $hourColumn++;
            }

            $sheet->setCellValueByColumnAndRow(35,$employeeColumnStartingPoint, "=SUM(C$employeeColumnStartingPoint:AG$employeeColumnStartingPoint)")
                ->getStyleByColumnAndRow(35,$employeeColumnStartingPoint)
                ->getAlignment()
                ->setVertical('center')->setHorizontal('center');

            $employeeColumnStartingPoint = $employeeColumnEndingPoint + 1;
            $employeeColumnEndingPoint = $employeeColumnStartingPoint + 2;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save("public/workHourSheet_$this->monthAndYear.xlsx");
    }
}
