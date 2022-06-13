<?php 	
	session_start();
	$_SESSION=array();
	$ulr = (parse_url($_SERVER['HTTP_REFERER']));
	var_dump($ulr["scheme"]."://".$ulr["host"]);
	header("Location: ".$ulr["scheme"]."://".$ulr["host"]);
?>