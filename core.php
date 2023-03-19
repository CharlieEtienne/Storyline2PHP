<?php
// Turn off all error reporting
error_reporting(0);

require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\IOFactory;
use JetBrains\PhpStorm\NoReturn;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class Storyline2PHP {

	public string $current_id;
	public array  $notes_text;
	public int    $id_col_index;
	public string $current_title;
	public int    $notes_col_index;

	public function __construct() {
		$this->current_id      = '';
		$this->notes_text      = ['Commentaires de la diapositive', 'Slide Notes', 'Foliennotizen', 'Notas de la diapositiva', '幻灯片备注' ];
		$this->id_col_index    = 0;
		$this->current_title   = '';
		$this->notes_col_index = 3;
	}

	/**
	 * Main function that runs the program
	 *
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function run(): void {
		// Start a session to store data across requests
		session_start();

		// If a file has been uploaded
		if ($_SERVER['REQUEST_METHOD'] == 'POST' && !empty($_FILES['file']['tmp_name'])) {

			// Define the allowed file extensions
			$allowed_extensions = array('docx');
			// Get the file extension of the uploaded file
			$file_extension = strtolower(pathinfo($_FILES['file']['name'], PATHINFO_EXTENSION));

			// If the file extension is not in the allowed list, throw an error
			if (!in_array($file_extension, $allowed_extensions)) {
				die('Error: The file must be in .docx format');
			}

			// If the file size is too large, throw an error
			if ($_FILES['file']['size'] > 10000000) {
				die('Error: The file is too large. Maximum size is 10 MB.');
			}

			// Load the uploaded file using the PhpWord library
			$phpWord  = IOFactory::load($_FILES['file']['tmp_name']);

			// Get all sections in the document
			$sections = $phpWord->getSections();

			if (empty($sections)){
				?>
				<div class="alert alert-danger" role="alert">
					<strong>Your file appears to be empty.</strong> <br>
					It's sometimes caused by a tiny bug in Storyline export feature, if you just exported the file and never modified it. Here's what you can try:
					<ol>
						<li>Open your file in Microsoft Word</li>
						<li>Save it (don't need to modify it)</li>
						<li>Close it</li>
						<li>Finally try to upload it here again</li>
					</ol>
				</div>
				<?php
				// exit();
			}

			// Initialize an empty array to store the data
			$data     = [];

			// Iterate over each section in the document
			foreach ($sections as $section) {
				// Get all elements in the section
				$elements = $section->getElements();
				foreach ($elements as $element) {
					// If the element is a TextRun, store its text as the current title
					if ($element instanceof TextRun) {
						$this->current_title = $this->getTextFromTextRun($element);
					}
					// If the element is a Table, iterate over its rows and store the notes data
					if ($element instanceof Table) {
						$rowData = $this->iterateOverRows($element);
						if (!empty($rowData)) {
							$data[] = $rowData;
						}
					}
				}
			}

			// Store the data in the session variable
			$_SESSION['data'] = $data;

			// Print the HTML table with the data
			$this->printHtmlTable($data);

		}
		// If the request method is POST and the export button has been clicked, then export the data to an Excel file.
		elseif ($_SERVER['REQUEST_METHOD'] == 'POST' && !empty($_POST['export'])) {
			$this->exportToExcel();
		}
		// If the request method is not POST or the export button has not been clicked, then render the upload form.
		else {
			$this->renderForm();
		}
	}

	/**
	*	This function iterates over each row in a table element and extracts the notes' data.
	*
	*	@param Table    $table    The table element to iterate over.
	*
	*	@return array An array containing the current title as the key and the notes text as the value.
    */
	private function iterateOverRows( Table $table ): array {
		$notes = "";
		foreach( $table->getRows() as $row ) {
			$cells = $row->getCells();
			foreach( $cells as $key => $cell ) {
				$els = $cell->getElements();
				foreach( $els as $e ) {
					$switchElements = $this->switchElements($e);
					if( in_array($switchElements, $this->notes_text) ) {
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

	/**
	 * This function receives an element from a Word document and switches it based on its type.
	 *
	 * @param mixed $element The element to be switched.
	 *
	 * @return array|bool|string The switched element, or false if the element's type is not supported.
	 */
	private function switchElements( mixed $element ): array|bool|string {
		if( $element instanceof TextRun ) {
			return $this->getTextFromTextRun($element);
		} elseif( $element instanceof Table ) {
			return $this->iterateOverRows($element);
		}
		return false;
	}

	/**
	* Extracts text from a TextRun element.
	 *
	* @param mixed $element the TextRun element to extract text from
	* @return string the extracted text from the TextRun element
	*/
	private function getTextFromTextRun( mixed $element ): string {
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
	 * Export data to Excel format and download.
	 *
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 * @return void
	 */
	#[NoReturn] public function exportToExcel(): void {

		// Create new Spreadsheet instance
		$spreadsheet = new Spreadsheet();
		$sheet = $spreadsheet->getActiveSheet();

		// Add column headers
		$sheet->setCellValue('A1', 'Slide Title');
		$sheet->setCellValue('B1', 'Notes');
		$sheet->getStyle('A1:B1')->getFont()->setBold( true );

		// Fill in the data
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

		// Format column styles
		$sheet->getStyle('A:B')->getAlignment()->setWrapText(true);
		$sheet->getStyle('A:B')->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
		$sheet->getColumnDimensionByColumn(1)->setAutoSize(false);
		$sheet->getColumnDimensionByColumn(1)->setWidth('40');
		$sheet->getColumnDimensionByColumn(2)->setAutoSize(false);
		$sheet->getColumnDimensionByColumn(2)->setWidth('120');


		// Download the file
		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
		$fileName = 'export';
		ob_end_clean();
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header("Content-Disposition: attachment; filename=\"$fileName.xlsx\"");
		$writer->save('php://output');
		exit();
	}


	/**
	 * This method prints an HTML table based on the data array provided.
	 *
	 * @param array    $data The data array containing slide titles and notes.
	 * @return void
	 */
	public function printHtmlTable(array $data): void {
		$i = 0;
		?>
		<h2 class="text-center">Extract notes from Storyline Word export file</h2>
		<div class="actions my-4 text-center">
			<form method="get" class="d-inline-block"><button class="btn btn-primary" type="submit" name="display_form">Try with another file</button></form>
			<form method="post" class="d-inline-block ms-3"><button class="btn btn-primary" type="submit" name="export" value="1">Export to Excel</button></form>
			<button class="d-inline-block ms-3 btn btn-primary copy--all">Copy to clipboard</button>
		</div>
		<div class="shadow-sm overflow-hidden p-4 rounded table-container">
			<table class="rounded table table-striped">
				<thead><tr><th>Slide title</th><th>Notes</th><th></th></tr></thead>
				<tbody>
				<?php foreach ($data as $rows): ?>
					<?php foreach ($rows as $id => $notes): ?>
						<?php $notes = nl2br($notes); ?>
						<tr>
							<td style="width: 25%"><strong><?= $id ?></strong></td>
							<td id="content--<?= $i ?>"><?= $notes ?></td>
							<td style="width: 80px"><button class="btn btn-dark copy--btn" data-clipboard-target="#content--<?= $i ?>" title="Copy to clipboard">Copy</button></td>
						</tr>
						<?php $i++; ?>
					<?php endforeach; ?>
				<?php endforeach; ?>
				</tbody>
			</table>
		</div>
		<div class="actions mt-4 text-center">
			<form method="get" class="d-inline-block"><button class="btn btn-primary" type="submit" name="display_form">Try with another file</button></form>
			<form method="post" class="d-inline-block ms-3"><button class="btn btn-primary" type="submit" name="export" value="1">Export to Excel</button></form>
			<button class="d-inline-block ms-3 btn btn-primary copy--all">Copy to clipboard</button>
		</div>
		<?php
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
		?>
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
				<input class="btn btn-lg btn-primary" type="submit" value="Get Slide Notes!">
			</div>
		</form>
		<?php
	}
}

$instance = new Storyline2PHP();
try {
	$instance->run();
} catch( \PhpOffice\PhpSpreadsheet\Writer\Exception $e ) {
	echo $e->getMessage();
}