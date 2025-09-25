<?php

namespace App\Http\Controllers\Web;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use Barryvdh\DomPDF\Facade\Pdf;
use Maatwebsite\Excel\Facades\Excel;
use App\Exports\LabelExport;
use App\Label;
use App\Increment;

class IndexController extends Controller
{
    /** 🔹 DASHBOARD */
    public function index(Request $req)
    {
        $user = auth()->user();

        $labels = Label::with('operator')->orderBy('id', 'desc');

        // kalau mau filter khusus operator
        // if(!$user->isAdmin()) {
        //     $labels->where('operator_id', $user->id);
        // }

        if ($req->filled('search') || $req->search === "0") {
            $labels->where('lot_number', 'like', '%' . $req->search . '%');
        }

        $labels = $labels->paginate(10);

        return view('web.dashboard.index', compact("labels"));
    }

    /** 🔹 Ambil next lot (AJAX hint) */
    public function getNextLot(Request $req)
    {
        $machine = $req->machine;
        $date = $req->date;

        if (!$machine || !$date) {
            return response()->json(['next_lot' => null]);
        }

        $next = $this->getIncrement($machine, $date);
        $pharse = str_pad($next, 3, '0', STR_PAD_LEFT);

        return response()->json(['next_lot' => $pharse]);
    }

    /** 🔹 PRINT SATU LABEL */
    public function print(Request $req)
    {
        $last_number = $this->getIncrement($req->machine_number, $req->shift_date);
        $pharse = $this->pharseLastNumber($last_number);

        $user = auth()->user();

        $label = new Label;
        $label->lot_number = $req->machine_number . date('ymd', strtotime($req->shift_date)) . $pharse;
        $label->formated_lot_number = $req->machine_number . "-" . date('y-m-d', strtotime($req->shift_date)) . "-" . $pharse;
        $label->size = $req->size;
        $label->length = $req->length;
        $label->weight = $req->weight;
        $label->shift_date = $req->shift_date;
        $label->shift = $req->shift;
        $label->machine_number = $req->machine_number;
        $label->pitch = $req->pitch;
        $label->direction = $req->direction;
        $label->visual = $req->visual;
        $label->remark = $req->remark;
        $label->bobin_no = $req->bobin_no;
        $label->operator_id = $user->id;
        $label->print_count = 1;
        $label->save();

        return view('export.label', ['label' => $label]);
    }

    public function dataLabel(Request $request)
    {
        $labels = Label::with('operator')
            ->when($request->search, function ($q) use ($request) {
                return $q->where('lot_number', 'like', "%{$request->search}%");
            })
            ->orderBy('id', 'desc')
            ->paginate(10);
            
        return view('web.label.index', compact('labels'));
    }

    public function edit($label)
    {
        $label = Label::find($label);
        if (!$label) {
            return redirect()->route("web.dashboard.index");
        }

        return view('web.dashboard.edit', compact("label"));
    }


    public function update(Request $req, $label)
    {
        $user = auth()->user();
        $label = Label::find($label);
        if (!$label) {
            return redirect()->route("web.dashboard.index");
        }

        $pharse = $req->lot_not;

        $label->lot_number = $req->machine_number . date('ymd', strtotime($req->shift_date)) . $pharse;
        $label->formated_lot_number = $req->machine_number . "-" . date('y-m-d', strtotime($req->shift_date)) . "-" . $pharse;
        $label->size = $req->size;
        $label->length = $req->length;
        $label->weight = $req->weight;
        $label->shift_date = $req->shift_date;
        $label->shift = $req->shift;
        $label->machine_number = $req->machine_number;
        $label->pitch = $req->pitch;
        $label->direction = $req->direction;
        $label->visual = $req->visual;
        $label->remark = $req->remark;
        $label->bobin_no = $req->bobin_no;
        $label->operator_id = $user->id;
        $label->save();

        $label->increment('print_count');

        return view('export.label', ['label' => $label]);
    }

    public function printSingle($id)
    {
    $label = Label::with('operator')->findOrFail($id);
    $label->increment('print_count');
    return view('export.label', compact('label'));
    }

    public function updateOnly(Request $req, $label)
    {
        $user = auth()->user();
        $label = Label::find($label);
        if (!$label) {
            return redirect()->route("web.dashboard.index");
        }

        $pharse = $req->lot_not ?? $this->pharseLastNumber(intval(substr($label->lot_number, -3)));

        $label->lot_number = $req->machine_number . date('ymd', strtotime($req->shift_date)) . $pharse;
        $label->formated_lot_number = $req->machine_number . "-" . date('y-m-d', strtotime($req->shift_date)) . "-" . $pharse;
        $label->size = $req->size;
        $label->length = $req->length;
        $label->weight = $req->weight;
        $label->shift_date = $req->shift_date;
        $label->shift = $req->shift;
        $label->machine_number = $req->machine_number;
        $label->pitch = $req->pitch;
        $label->direction = $req->direction;
        $label->visual = $req->visual;
        $label->remark = $req->remark;
        $label->bobin_no = $req->bobin_no;
        $label->operator_id = $user->id;
        $label->save();

        return redirect()->route("web.label.index");
    }


    public function delete($label)
    {
        $label = Label::find($label);
        if ($label) {
            $label->delete();
        }

        return redirect()->route('web.label.index');
    }

    public function exportExcel(Request $request)
    {
        $startDate = $request->start_date;
        $endDate = $request->end_date;

        return Excel::download(
            new LabelExport($startDate, $endDate),
            'data_label_.xlsx'
        );
    }

    public function exportPDF(Request $request)
    {
        ini_set('max_execution_time', 120);
        ini_set('memory_limit', '1024M');

        $query = Label::with('operator');

        // 🔹 Kalau ada filter tanggal
        if ($request->has('start_date') && $request->has('end_date')) {
            $query->whereBetween('shift_date', [
                $request->start_date,
                $request->end_date
            ]);
        }

        $labels = $query->orderBy('id', 'desc')->limit(1024)->get();

        $pdf = Pdf::loadView('export.label_pdf', compact('labels'))
            ->setPaper('a4', 'landscape');

        return $pdf->download('data_label.pdf');
    }

    public function printView(Request $request)
    {
        $start = $request->input('start_date');
        $end = $request->input('end_date');

        $query = Label::with('operator');

        if ($start && $end) {
            $query->whereBetween('shift_date', [$start, $end]);
        }

        $labels = $query->orderBy('created_at', 'desc')->get();

        return view('web.label.print', compact('labels', 'start', 'end'));
    }


    /** Helpers */
    private function getIncrement($machine_number, $shift_date)
    {
        $lastLabel = Label::where('machine_number', $machine_number)
            ->whereDate('shift_date', $shift_date)
            ->orderBy('lot_number', 'desc')
            ->first();

        if (!$lastLabel) {
            return 1;
        }

        $lastNumber = intval(substr($lastLabel->lot_number, -3));
        return $lastNumber + 1;
    }

    private function pharseLastNumber($number)
    {
        return str_pad($number, 3, '0', STR_PAD_LEFT);
    }
}

