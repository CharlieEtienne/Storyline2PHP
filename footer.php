</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js" integrity="sha384-w76AqPfDkMBDXo30jS1Sgez6pr3x5MlQ1ZAGC+nuZB+EYdgRZgiwxhTBTkF7CXvN" crossorigin="anonymous"></script>
<script src="https://unpkg.com/filepond@^4/dist/filepond.js"></script>
<script>
	// Get a file input reference
	const input = document.querySelector('input[type="file"]');

	// Create a FilePond instance
	FilePond.create(input, {
		storeAsFile: true,
	});
</script>
</body>
</html>