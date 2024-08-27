<?php

namespace App\Exports;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class ReserveExport implements WithMultipleSheets
{
    protected $data;
    protected $sel_doc;

    public function __construct(array $data, array $sel_doc)
    {
        $this->data = $data;
        $this->sel_doc = $sel_doc;

    }

    public function sheets(): array
    {
        // dd($this->data);
        $sheets = [];

        $sheets[] = new namelist($this->data);
        $sheets[] = new booking($this->data);
        $sheets[] = new Condition($this->data);
        $sheets[] = new Technician_Notice($this->data);
        $sheets[] = new Purchase_sale($this->data);
        $sheets[] = new Customer_check($this->data);
        $sheets[] = new Mechanic_check($this->data);

        return $sheets;
    }
}

class namelist implements FromView, WithTitle, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.namelist', [
            'data' => $this->data,
        ]);
    }

    public function title(): string
    {
        return 'namelist';
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:CA100')->getFont()->setName('Tahoma');
            },
        ];
    }
}

class booking implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.booking', [
            'data' => $this->data,
        ]);
    }
    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setPath(public_path('/images/logo/logo_kp.png'));
        $drawing->setHeight(90);
        $drawing->setCoordinates('B1');

        return $drawing;
    }

    public function title(): string
    {
        return 'ใบจอง';
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Angsana New');
                $event->sheet->getStyle('B11:F11')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B7:B11')->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B7:F7')->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('F7:F11')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
            },
        ];
    }
}

class Condition implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.Condition', [
            'data' => $this->data,
        ]);
    }

    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setPath(public_path('/images/logo/logo_kp.png'));
        $drawing->setHeight(60);
        $drawing->setCoordinates('B1');
        $drawing->setOffsetX(20);
        $drawing->setOffsetY(10);

        return $drawing;
    }

    public function title(): string
    {
        return 'เงื่อนไข';
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Angsana New');
                $event->sheet->getStyle('A4:L4')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('I8:K8')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('I9:K9')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('D11:L11')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('D12:L12')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('D13:L13')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('D14:L14')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('D15:F15')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('I15:L15')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
            },
        ];
    }
}

class Technician_Notice implements FromView, WithTitle, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.Technician_Notice', [
            'data' => $this->data,
        ]);
    }

    public function title(): string
    {
        return 'ใบแจ้งช่าง';
    }
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Angsana New');
                $event->sheet->getStyle('B2:G2')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('I2:M2')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B5:M5')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B6:M6')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B8:M8')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H10:M10')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H18:M18')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H19:M19')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H24:M24')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H25:M25')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H29:M29')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H30:M30')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A3:A8')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('M3:M8')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('G3:G5')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('H3:H5')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
            },
        ];
    }
}

class Purchase_sale implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.Purchase_sale', [
            'data' => $this->data,
        ]);
    }

    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setPath(public_path('/images/logo/logo_kp.png'));
        $drawing->setHeight(90);
        $drawing->setCoordinates('B2');

        return $drawing;
    }

    public function title(): string
    {
        return 'สัญญาซื้อ-ขาย';
    }
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Angsana New');
                $event->sheet->getStyle('B34:L34')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B41:L41')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('B42:L42')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);

                $event->sheet->getStyle('A35:A42')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('L35:L42')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);

            },
        ];
    }
}

class Customer_check implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.Customer_check', [
            'data' => $this->data,
        ]);
    }
    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setPath(public_path('/images/logo/logo_kp2.png'));
        $drawing->setHeight(60);
        $drawing->setCoordinates('E1');
        $drawing->setOffsetx(20);

        return $drawing;
    }

    public function title(): string
    {
        return 'ใบเช็ครถลูกค้า';
    }
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Angsana New');
                $event->sheet->getStyle('A10:L10')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A1:C1')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('A2:C2')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('A2')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);

                $event->sheet->getStyle('L10')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
            },
        ];
    }
}

class Mechanic_check implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.Mechanic_check', [
            'data' => $this->data,
        ]);
    }
    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setPath(public_path('/images/logo/logo_kp2.png'));
        $drawing->setHeight(60);
        $drawing->setCoordinates('E1');
        $drawing->setOffsetx(20);

        return $drawing;
    }


    public function title(): string
    {
        return 'ใบเช็ครถของช่าง';
    }
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Angsana New');
                $event->sheet->getStyle('A11:L11')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A12:L12')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A22:L22')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A23:L23')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A32:L32')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A33:L33')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A40:L40')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A42:L42')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A43:L43')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A54:L54')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('A55:L55')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);

                $event->sheet->getStyle('L10:L12')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('L23')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('L33')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('L55')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);
                $event->sheet->getStyle('L41:L43')->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM);

                $event->sheet->getStyle('B63:L63')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('A64:L64')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('A65:L65')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('A66:L66')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);
                $event->sheet->getStyle('A67:L67')->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR);

            },
        ];
    }
}
class Delivery_note implements FromView, WithTitle, WithDrawings, WithEvents
{
    protected $data;

    public function __construct(array $data)
    {
        $this->data = $data;
    }

    public function view(): View
    {

        // dd($aggregatedData);
        return view('export.Delivery_note', [
            'data' => $this->data,
        ]);
    }
    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setPath(public_path('/images/logo/logo_kp_add.png'));
        $drawing->setHeight(120);
        $drawing->setCoordinates('B2');

        return $drawing;
    }


    public function title(): string
    {
        return 'ใบเช็ครถของช่าง';
    }
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->getStyle('A1:S100')->getFont()->setName('Tahoma');

            },
        ];
    }
}
