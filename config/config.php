<?php 

	ob_start(); // Start output buffering
	$temp = ob_get_contents(); // Store buffer in variable
	require_once $_SERVER['DOCUMENT_ROOT']."/Framework/Framework.php";
	require_once  $_SERVER['DOCUMENT_ROOT']."/api/classes.php";
	ob_end_clean(); // End buffering and clean up

	Connection::initialize("unificar_corpo");
	
?>