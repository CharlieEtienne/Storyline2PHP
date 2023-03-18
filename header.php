<!doctype html>
<html lang="en" data-bs-theme="dark">
<head>
	<meta charset="UTF-8">
	<meta name="viewport"
	      content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
	<meta http-equiv="X-UA-Compatible" content="ie=edge">
	<title>Storyline2PHP</title>
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-GLhlTQ8iRABdZLl6O3oVMWSktQOp6b7In1Zl3/Jr59b6EGGoI1aFkw7cmDA6j6gD" crossorigin="anonymous">
	<link href="https://unpkg.com/filepond@^4/dist/filepond.css" rel="stylesheet" />
	<link href="https://fonts.bunny.net/css2?family=Nunito:wght@400;600;700&amp;display=swap" rel="stylesheet">
	<style>

        [data-bs-theme=dark] {
            --bs-body-color: hsl(229 8% 72%);
            --bs-body-bg: hsl(229 40% 19%);
        }
        .btn-primary {
            --bs-: hsl(250, 90%, 66%);
            --bs-btn-bg: hsl(250, 90%, 66%);
            --bs-btn-border-color: hsl(250, 90%, 66%);
            --bs-btn-hover-bg: hsl(250, 71%, 55%);
            --bs-btn-hover-border-color: hsl(250, 81%, 52%);
            --bs-btn-active-bg: hsl(250, 81%, 52%);
            --bs-btn-active-border-color: hsl(250, 80%, 49%);
            --bs-btn-disabled-bg: hsl(250, 90%, 66%);
            --bs-btn-disabled-border-color: hsl(250, 90%, 66%);
		}
        .table-container {
            background: hsl(229, 40%, 15%);
        }
        .filepond--root {
            height: 300px;
	        border: none !important;
        }
        .filepond--root .filepond--drop-label {
            min-height: 300px;
        }
        .filepond--panel-root {
            background: var(--bs-form-control-bg);
            border: 3px dashed hsl(229 15% 32%);
        }
        .filepond--drop-label {
            color: var(--bs-body-color);
        }
        .filepond--credits {
            color: white !important;
            opacity: 0.3 !important;
        }
	</style>
</head>
<body>
<div class="container py-5">
