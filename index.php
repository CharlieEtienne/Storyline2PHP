<?php
require_once 'vendor/autoload.php';

include_once 'header.php';

use PhpOffice\PhpWord\IOFactory;
use JetBrains\PhpStorm\NoReturn;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class Storyline2PHP {

	public string $filepath;
	public string $current_id;
	public string $notes_text;
	public int    $id_col_index;
	public string $current_title;
	public int    $notes_col_index;

	public function __construct() {
		$this->filepath        = 'tableau.docx';
		$this->current_id      = '';
		$this->notes_text      = 'Commentaires de la diapositive';
		$this->id_col_index    = 0;
		$this->current_title   = '';
		$this->notes_col_index = 3;
	}

	/**
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function run(): void {
		session_start();
		if ($_SERVER['REQUEST_METHOD'] == 'POST' && !empty($_FILES['file']['tmp_name'])) {

			$allowed_extensions = array('docx');
			$file_extension = strtolower(pathinfo($_FILES['file']['name'], PATHINFO_EXTENSION));

			if (!in_array($file_extension, $allowed_extensions)) {
				die('Error: The file must be in .docx format');
			}

			if ($_FILES['file']['size'] > 10000000) {
				die('Error: The file is too large. Maximum size is 10 MB.');
			}

			$phpWord  = IOFactory::load($_FILES['file']['tmp_name']);
			$sections = $phpWord->getSections();
			$data     = [];

			foreach ($sections as $section) {
				$elements = $section->getElements();
				foreach ($elements as $element) {
					if ($element instanceof TextRun) {
						$this->current_title = $this->getTextFromTextRun($element);
					}
					if ($element instanceof Table) {
						$rowData = $this->iterateOverRows($element);
						if (!empty($rowData)) {
							$data[] = $rowData;
						}
					}
				}
			}

			$_SESSION['data'] = $data;
			$this->printHtmlTable($data);
		}
		elseif ($_SERVER['REQUEST_METHOD'] == 'POST' && !empty($_POST['export'])) {
			// var_dump($_SESSION['data']);
			$this->exportToExcel();

		} else {
			$this->renderForm();
		}
	}

	private function iterateOverRows( $table ): array {
		$notes = "";
		foreach( $table->getRows() as $row ) {
			$cells = $row->getCells();
			foreach( $cells as $key => $cell ) {
				$els = $cell->getElements();
				foreach( $els as $e ) {
					$switchElements = $this->switchElements($e);
					if( $switchElements === $this->notes_text ) {
						$this->current_id = $this->getNotesId($cells[ $this->id_col_index ]);
					}
				}
				if( $this->getNotesId($cells[ $this->id_col_index ]) === $this->current_id && $key === $this->notes_col_index ) {
					foreach( $els as $e ) {
						$notes .= $this->switchElements($e) . "\n";
					}
				}
			}
		}
		if( !empty($notes) ) {
			return [ $this->current_title => $notes ];
		}
		return [];
	}

	private function switchElements( $element ): array|bool|string {
		if( $element instanceof TextRun ) {
			return $this->getTextFromTextRun($element);
		} elseif( $element instanceof Table ) {
			return $this->iterateOverRows($element);
		}
		return false;
	}

	private function getTextFromTextRun( $element ): string {
		$text = "";
		for( $index = 0; $index < $element->countElements(); $index++ ) {
			$textRunElement = $element->getElement($index);
			if( $textRunElement instanceof Text || $textRunElement instanceof TextRun ) {
				$text .= $textRunElement->getText();
			}
		}
		return $text;
	}

	/**
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	#[NoReturn] public function exportToExcel(): void {
		$spreadsheet = new Spreadsheet();
		$sheet = $spreadsheet->getActiveSheet();

		// Ajouter les en-têtes de colonne
		$sheet->setCellValue('A1', 'Slide Title');
		$sheet->setCellValue('B1', 'Notes');
		$sheet->getStyle('A1:B1')->getFont()->setBold( true );

		// Remplir les données
		$row = 2;
		$colored = false;
		foreach ($_SESSION['data'] as $rows) {
			foreach ($rows as $id => $notes) {
				$sheet->setCellValue('A' . $row, $id);
				$sheet->setCellValue('B' . $row, $notes);
				if($colored){
					$sheet->getStyle('A' . $row . ':B' . $row)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('f2f2f2');
				}
				$row++;
				$colored = !$colored;
			}
		}

		$sheet->getStyle('A:B')->getAlignment()->setWrapText(true);
		$sheet->getStyle('A:B')->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
		$sheet->getColumnDimensionByColumn(1)->setAutoSize(false);
		$sheet->getColumnDimensionByColumn(1)->setWidth('40');
		$sheet->getColumnDimensionByColumn(2)->setAutoSize(false);
		$sheet->getColumnDimensionByColumn(2)->setWidth('120');


		// Télécharger le fichier
		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
		$fileName = 'export';
		ob_end_clean();
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header("Content-Disposition: attachment; filename=\"$fileName.xlsx\"");
		$writer->save('php://output');
		exit();
	}


	/**
	 * @param array    $data
	 *
	 * @return void
	 */
	public function printHtmlTable( array $data ): void {
		echo "<div class='actions mb-4 text-center'>";
		echo '<form method="get" class="d-inline-block"><button class="btn btn-primary" type="submit" name="display_form">Try with another file</button></form>';
		echo '<form method="post" class="d-inline-block ms-3"><button class="btn btn-primary" type="submit" name="export" value="1">Export to Excel</button></form>';
		echo "</div>";

		echo "<div class='shadow-sm overflow-hidden p-4 rounded table-container'>";
		echo "<table class='rounded table table-striped'>";
		echo "<thead><tr><th>Slide title</th><th>Notes</th></tr></thead>";
		echo "<tbody>";
		foreach( $data as $rows ) {
			foreach( $rows as $id => $notes ) {
				$notes = nl2br($notes);
				echo "<tr><td style='width: 25%'><strong>$id</strong></td><td>$notes</td></tr>";
			}
		}
		echo "</tbody>";
		echo "</table>";
		echo "</div>";

		echo "<div class='actions mt-4 text-center'>";
		echo '<form method="get" class="d-inline-block"><button class="btn btn-primary" type="submit" name="display_form">Try with another file</button></form>';
		echo '<form method="post" class="d-inline-block ms-3"><button class="btn btn-primary" type="submit" name="export" value="1">Export to Excel</button></form>';
		echo "</div>";
	}

	/**
	 * @param $cells
	 *
	 * @return array|false|string|string[]
	 */
	public function getNotesId( $cells ): string|array|false {
		return $this->switchElements($cells->getElements()[ 0 ]);
	}

	public function renderForm(): void {
		echo <<<HTML
		<header class="text-center">
			<h2>Extract notes from Storyline Word export file</h2>
			<p>In Articulate Storyline, go to File > Translation > Export<br>
			Drop the .docx file to extract notes only!</p>
		</header>
		<form class="mt-5" method="post" enctype="multipart/form-data">
			<div class="text-center">
				<label for="file">Select a Word file (.docx)</label>
				<input class="form-control my-3" type="file" id="file" name="file" accept=".docx" required>
			</div>
			<div class="text-center">
				<input class="btn btn-lg btn-primary" type="submit" value="Send file">
			</div>
		</form>
		HTML;
	}
}

$instance = new Storyline2PHP();
try {
	$instance->run();
} catch( \PhpOffice\PhpSpreadsheet\Writer\Exception $e ) {
	echo $e->getMessage();
}

include_once 'footer.php';