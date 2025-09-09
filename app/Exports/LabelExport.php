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

class LabelExport implements FromCollection, WithHeadings, WithMapping, ShouldAutoSize, WithColumnFormatting, WithStyles
{
    protected $startDate;
    protected $endDate;
    protected $rowNumber = 0;

    public function __construct($startDate = null, $endDate = null)
    {
        $this->startDate = $startDate;
        $this->endDate   = $endDate;
    }

    public function collection()
    {
        $query = Label::select(
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

        return $query->get();
    }

    public function headings(): array
    {
        return [
            'No',
            'ID',
            'Lot Number',
            'Formatted Lot Number',
            'Size',
            'Length (m)',
            'Weight (kg)',
            'Shift Date',
            'Shift',
            'Machine Number',
            'Pitch (mm)',
            'Direction',
            'Visual',
            'Remark',
            'Bobin No',
            'Operator Name',
            'Created At'
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
            Carbon::parse($row->shift_date)->format('d-m-Y'),
            $row->shift,
            $row->machine_number,
            $row->pitch,
            $row->direction,
            $row->visual,
            $row->remark,
            $row->bobin_no,
            $row->operator_name,
            // $row->created_at,
            Carbon::parse($row->created_at)->format('d-m-Y H:i'),
        ];
    }

    public function columnFormats(): array
    {
        return [
            'H' => NumberFormat::FORMAT_DATE_DDMMYYYY,      // Shift Date
            'Q' => NumberFormat::FORMAT_DATE_DATETIME,      // Created At
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
