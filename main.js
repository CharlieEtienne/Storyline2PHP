// noinspection JSUnresolvedVariable,JSUnresolvedFunction

// Get a file input reference
const input = document.querySelector( 'input[type="file"]' );

// Create a FilePond instance
FilePond.create( input, {
	storeAsFile: true,
} );

const table         = document.querySelector( 'table' );
const copyBtnSel    = '.copy--btn';
const copyAllBtnSel = '.copy--all';
let tableCopy;

function cloneTable() {
	// Create a temporary copy of the table
	tableCopy = table.cloneNode( true );

	// Remove the third column from the table copy
	const colIndexToRemove = 2;
	const thToRemove       = tableCopy.querySelector( `th:nth-child(${colIndexToRemove + 1})` );
	const tdToRemove       = tableCopy.querySelectorAll( `td:nth-child(${colIndexToRemove + 1})` );
	thToRemove.remove();
	tdToRemove.forEach( td => td.remove() );

	// Hide the table copy
	tableCopy.style.position = 'fixed';
	tableCopy.style.left     = '-10000px';

	// Add the table copy to the document body
	document.body.appendChild( tableCopy );
}

cloneTable();

/**
 * Initialize tooltip for all matching selectors
 *
 * @param Sel
 */
function tooltipInit(Sel) {
	document.querySelectorAll( Sel ).forEach( e => {
		const tooltip = new bootstrap.Tooltip( e, {trigger: 'click'} );
		e.setAttribute( 'data-bs-original-title', 'Copied!' );
		e.addEventListener( 'mouseleave', () => { tooltip.hide() } )
	} );
}

// Init Tooltips
tooltipInit( copyBtnSel );
tooltipInit( copyAllBtnSel );

// Clipboard
const clipboard    = new ClipboardJS( copyBtnSel );
const clipboardAll = new ClipboardJS( copyAllBtnSel, {
	target: () => tableCopy
} );

/**
 * Initialize ClipboardJS success callback for all instances
 *
 * @param clipboardInstance
 */
function clipboardInit(clipboardInstance) {
	clipboardInstance.on( 'success', e => {
		const tooltip = bootstrap.Tooltip.getInstance( e.trigger );
		tooltip.show();
		e.clearSelection()
	} );
}

// Init ClipboardJS instances
clipboardInit( clipboard );
clipboardInit( clipboardAll );