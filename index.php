<?php
require_once 'vendor/autoload.php';

include_once 'header.php';

use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Element\TextRun;

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

	public function run(): void {
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

			$this->printHtmlTable($data);
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
	 * @param array    $data
	 *
	 * @return void
	 */
	public function printHtmlTable( array $data ): void {
		echo "<div class='table-container rounded overflow-hidden'>";
		echo "<table class='table table-striped table-bordered rounded'>";
		echo "<tr><th>Slide title</th><th>Notes</th></tr>";
		foreach( $data as $rows ) {
			foreach( $rows as $id => $notes ) {
				$notes = nl2br($notes);
				echo "<tr><td>$id</td><td>$notes</td></tr>";
			}
		}
		echo "</table>";
		echo "</div>";
		echo '<form method="get"><button class="btn btn-lg btn-primary" type="submit" name="display_form">Try with another file</button></form>';
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
		<form method="post" enctype="multipart/form-data">
			<div>
				<label for="file">Select a Word file (.docx) :</label>
				<input class="form-control" type="file" id="file" name="file" accept=".docx" required>
			</div>
			<br>
			<div>
				<input class="btn btn-lg btn-primary" type="submit" value="Send file">
			</div>
		</form>
		HTML;
	}
}

$instance = new Storyline2PHP();
$instance->run();

include_once 'footer.php';