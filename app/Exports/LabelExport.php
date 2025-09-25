<?php

namespace App\Exports;

use App\Label;
use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithColumnFormatting;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\FromQuery;
use PhpOffice\PhpSpreadsheet\Shared\Date;

class LabelExport implements FromQuery, WithHeadings, WithMapping, ShouldAutoSize, WithColumnFormatting, WithStyles
{
    protected $startDate;
    protected $endDate;
    protected $rowNumber = 0;

    public function __construct($startDate = null, $endDate = null)
    {
        $this->startDate = $startDate;
        $this->endDate   = $endDate;
    }

    public function query()
    {
        $query = Label::query()
        ->select(
            'labels.id',
            'labels.lot_number',
            'labels.formated_lot_number',
            'labels.size',
            'labels.length',
            'labels.weight',
            'labels.shift_date',
            'labels.shift',
            'labels.machine_number',
            'labels.pitch',
            'labels.direction',
            'labels.visual',
            'labels.remark',
            'labels.bobin_no',
            'users.name as operator_name',
            'labels.created_at'
        )
        ->leftJoin('users', 'labels.operator_id', '=', 'users.id')
        ->orderBy('labels.id', 'desc');

        if ($this->startDate && $this->endDate) {
            $query->whereBetween('labels.shift_date', [$this->startDate, $this->endDate]);
        }

        return $query;
    }

    public function headings(): array
    {
        return [
            'No',
            'ID',
            'Lot Number',
            'Formatted Lot Number',
            'Type/Size',
            'Length (m)',
            'Weight (kg)',
            'Date',
            'Shift',
            'Machine Number',
            'Pitch (mm)',
            'Direction',
            'Visual',
            'Remark',
            'Operator Name',
        ];
    }

    public function map($row): array
    {
        $this->rowNumber++;

        return [
            $this->rowNumber,
            $row->id,
            "'" . $row->lot_number,
            "" . $row->formated_lot_number,
            $row->size,
            $row->length,
            $row->weight,
            // $row->shift_date,
            Date::PHPToExcel(Carbon::parse($row->shift_date)),
            $row->shift,
            $row->machine_number,
            $row->pitch,
            $row->direction,
            $row->visual,
            $row->remark,
            $row->operator_name,
            // $row->created_at,
        ];
    }

    public function columnFormats(): array
    {
        return [
            'C' => NumberFormat::FORMAT_TEXT,               // Lot Number
            'D' => NumberFormat::FORMAT_TEXT,               // Formatted Lot Number
            'F' => NumberFormat::FORMAT_NUMBER_00,         // Length (m)
            'G' => NumberFormat::FORMAT_NUMBER_00,         // Weight (kg)
            'H' => NumberFormat::FORMAT_DATE_DDMMYYYY,      // Date
            'K' => NumberFormat::FORMAT_NUMBER_00,         // Pitch (mm)
        ];
    }

    public function styles(Worksheet $sheet)
    {
        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();

        // Header bold dan center
        $sheet->getStyle('A1:'.$highestColumn.'1')->applyFromArray([
            'font' => ['bold' => true],
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ]);

        // Semua data pakai border tipis
        $sheet->getStyle('A2:'.$highestColumn.$highestRow)->applyFromArray([
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ]);
    }
}
